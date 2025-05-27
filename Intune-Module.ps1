function New-TenantIntune {
    Write-LogMessage -Message "Starting Intune configuration..." -Type Info
    
    # COMPLETE module reset to match working user script exactly
    try {
        # Remove ALL Graph modules first to avoid conflicts
        Write-LogMessage -Message "Clearing all Graph modules to prevent conflicts..." -Type Info
        Get-Module Microsoft.Graph* | Remove-Module -Force -ErrorAction SilentlyContinue
        
        # Disconnect any existing sessions
        try {
            Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        }
        catch {
            # Ignore disconnect errors
        }
        
        # Force load ONLY the exact modules needed for Intune in exact order
        $intuneModules = @(
            'Microsoft.Graph.DeviceManagement',
            'Microsoft.Graph.Groups', 
            'Microsoft.Graph.Identity.DirectoryManagement'
        )
        
        Write-LogMessage -Message "Loading ONLY Intune modules in exact order..." -Type Info
        foreach ($module in $intuneModules) {
            try {
                # Remove any existing version first
                Get-Module $module | Remove-Module -Force -ErrorAction SilentlyContinue
                
                # Import fresh
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
        
        Write-LogMessage -Message "Connecting to Microsoft Graph with Intune scopes only..." -Type Info
        Connect-MgGraph -Scopes $intuneScopes -ErrorAction Stop | Out-Null
        
        $context = Get-MgContext
        Write-LogMessage -Message "Connected to Microsoft Graph as $($context.Account)" -Type Success
        Write-LogMessage -Message "Active scopes: $($context.Scopes -join ', ')" -Type Info -LogOnly
        
        # Verify required scopes are present
        $missingScopes = @()
        foreach ($scope in $intuneScopes) {
            if ($context.Scopes -notcontains $scope) {
                $missingScopes += $scope
            }
        }
        
        if ($missingScopes.Count -gt 0) {
            Write-LogMessage -Message "Missing required scopes: $($missingScopes -join ', ')" -Type Error
            Write-LogMessage -Message "Please reconnect with proper permissions" -Type Error
            return $false
        }
        
        Write-LogMessage -Message "Graph connection verified with required scopes" -Type Success
        
        # Create WindowsAutoPilot dynamic group first
        $autopilotGroup = New-WindowsAutoPilotGroup
        if (-not $autopilotGroup) {
            Write-LogMessage -Message "Failed to create WindowsAutoPilot group" -Type Warning
        }
        
        # Create update rings and device groups
        $updateRings = New-WindowsUpdateRings
        if (-not $updateRings) {
            Write-LogMessage -Message "Failed to create Windows Update rings" -Type Warning
        }
        
        # Enable LAPS prerequisite
        $lapsEnabled = Enable-WindowsLAPS
        if (-not $lapsEnabled) {
            Write-LogMessage -Message "LAPS enablement failed - LAPS policies may not work correctly" -Type Warning
        }
        
        # Create all configuration policies
        $policies = @()
        
        # Core security policies
        $policies += New-DefenderPolicy
        $policies += New-DefenderAntivirusPolicy  
        $policies += New-FirewallPolicy
        $policies += New-TamperProtectionPolicy
        $policies += New-EDRPolicy
        
        # BitLocker encryption
        $policies += New-BitLockerPolicy
        
        # LAPS (requires prerequisite)
        $policies += New-LAPSPolicy
        
        # Device configuration
        $policies += New-OneDrivePolicy
        $policies += New-PowerOptionsPolicy
        $policies += New-AdminAccountPolicy
        $policies += New-UnenrollmentPolicy
        
        # Application policies
        $policies += New-EdgePolicies
        $policies += New-OfficePolicies
        $policies += New-OutlookPolicy
        
        # Assign policies to device groups
        $deviceGroups = @("WindowsDeviceRing0", "WindowsDeviceRing1", "WindowsDeviceRing2")
        foreach ($policy in $policies) {
            if ($policy -and $policy.id) {
                Assign-PolicyToGroups -PolicyId $policy.id -GroupNames $deviceGroups
            }
        }
        
        Write-LogMessage -Message "Intune configuration completed successfully" -Type Success
        return $true
    }
    catch {
        Write-LogMessage -Message "Error in Intune configuration - $($_.Exception.Message)" -Type Error
        return $false
    }
}

# === Group Creation Functions ===

function New-WindowsAutoPilotGroup {
    Write-LogMessage -Message "Creating WindowsAutoPilot dynamic group..." -Type Info
    
    try {
        # Check if group already exists
        $existingGroup = Get-MgGroup -Filter "displayName eq 'WindowsAutoPilot'" -ErrorAction SilentlyContinue
        
        if ($existingGroup) {
            Write-LogMessage -Message "WindowsAutoPilot group already exists" -Type Warning
            return $existingGroup
        }
        
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
        return $result
    }
    catch {
        Write-LogMessage -Message "Failed to create WindowsAutoPilot group - $($_.Exception.Message)" -Type Error
        return $null
    }
}

function New-WindowsUpdateRings {
    Write-LogMessage -Message "Creating Windows Update rings and device groups..." -Type Info
    
    try {
        # Create device groups first
        $deviceGroups = @(
            @{ Name = "WindowsDeviceRing0"; Description = "Windows devices - Pilot ring (0/0 day deferrals)" },
            @{ Name = "WindowsDeviceRing1"; Description = "Windows devices - UAT ring (2/7 day deferrals)" },
            @{ Name = "WindowsDeviceRing2"; Description = "Windows devices - Production ring (7/7 day deferrals)" }
        )
        
        foreach ($group in $deviceGroups) {
            $existingGroup = Get-MgGroup -Filter "displayName eq '$($group.Name)'" -ErrorAction SilentlyContinue
            
            if (-not $existingGroup) {
                $groupBody = @{
                    displayName = $group.Name
                    description = $group.Description
                    mailEnabled = $false
                    mailNickname = $group.Name
                    securityEnabled = $true
                }
                
                $newGroup = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/groups" -Body $groupBody
                Write-LogMessage -Message "Created device group: $($group.Name)" -Type Success
                
                # Store group ID for later use
                $script:TenantState.CreatedGroups[$group.Name] = $newGroup.id
            }
            else {
                Write-LogMessage -Message "Device group already exists: $($group.Name)" -Type Warning
                $script:TenantState.CreatedGroups[$group.Name] = $existingGroup.Id
            }
        }
        
        # Create update rings
        $updateRings = @(
            @{ 
                Name = "RING0 - Pilot"; 
                Description = "Pilot testers for latest feature/quality updates"
                QualityDeferral = 0; 
                FeatureDeferral = 0;
                GroupName = "WindowsDeviceRing0"
            },
            @{ 
                Name = "RING1 - UAT"; 
                Description = "UAT testers for latest feature/quality updates"
                QualityDeferral = 2; 
                FeatureDeferral = 7;
                GroupName = "WindowsDeviceRing1"
            },
            @{ 
                Name = "RING2 - Production"; 
                Description = "Production devices with standard deferrals"
                QualityDeferral = 7; 
                FeatureDeferral = 7;
                GroupName = "WindowsDeviceRing2"
            }
        )
        
        foreach ($ring in $updateRings) {
            $ringBody = @{
                displayName = $ring.Name
                description = $ring.Description
                qualityUpdatesDeferralPeriodInDays = $ring.QualityDeferral
                featureUpdatesDeferralPeriodInDays = $ring.FeatureDeferral
                automaticUpdateMode = "autoInstallAtMaintenanceTime"
                businessReadyUpdatesOnly = "businessReadyOnly"
                skipChecksBeforeRestart = $false
                pauseFeatureUpdates = $false
                pauseQualityUpdates = $false
            }
            
            try {
                $result = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/windowsUpdateForBusinessRings" -Body $ringBody
                Write-LogMessage -Message "Created update ring: $($ring.Name)" -Type Success
                
                # Assign to corresponding device group
                if ($script:TenantState.CreatedGroups.ContainsKey($ring.GroupName)) {
                    $groupId = $script:TenantState.CreatedGroups[$ring.GroupName]
                    $assignmentBody = @{
                        assignments = @(
                            @{
                                target = @{
                                    "@odata.type" = "#microsoft.graph.groupAssignmentTarget"
                                    groupId = $groupId
                                }
                            }
                        )
                    }
                    
                    Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/windowsUpdateForBusinessRings/$($result.id)/assignments" -Body $assignmentBody
                    Write-LogMessage -Message "Assigned update ring $($ring.Name) to group $($ring.GroupName)" -Type Success
                }
            }
            catch {
                Write-LogMessage -Message "Failed to create update ring $($ring.Name) - $($_.Exception.Message)" -Type Error
            }
        }
        
        return $true
    }
    catch {
        Write-LogMessage -Message "Failed to create Windows Update rings - $($_.Exception.Message)" -Type Error
        return $false
    }
}

function Enable-WindowsLAPS {
    Write-LogMessage -Message "Checking Windows LAPS prerequisite..." -Type Info
    
    try {
        # Check if LAPS is already enabled
        $lapsSettings = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/policies/deviceRegistrationPolicy" -ErrorAction SilentlyContinue
        
        if ($lapsSettings -and $lapsSettings.localAdminPassword -and $lapsSettings.localAdminPassword.isEnabled) {
            Write-LogMessage -Message "Windows LAPS is already enabled" -Type Info
            return $true
        }
        
        # Enable Windows LAPS
        $body = @{
            localAdminPassword = @{
                isEnabled = $true
            }
        }
        
        Invoke-MgGraphRequest -Method PATCH -Uri "https://graph.microsoft.com/beta/policies/deviceRegistrationPolicy" -Body $body
        Write-LogMessage -Message "Windows LAPS has been enabled" -Type Success
        return $true
    }
    catch {
        Write-LogMessage -Message "Failed to enable Windows LAPS - $($_.Exception.Message)" -Type Error
        return $false
    }
}

# === Policy Creation Functions ===

function New-DefenderPolicy {
    Write-LogMessage -Message "Creating Defender Endpoint Security policy..." -Type Info
    
    try {
        $body = @{
            name = "Defender Configuration"
            description = "Microsoft Defender security configuration"
            platforms = "windows10"
            technologies = "mdm"
            templateReference = @{
                templateId = ""
                templateFamily = "none"
            }
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
                },
                @{
                    id = "1"
                    settingInstance = @{
                        "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                        settingDefinitionId = "device_vendor_msft_policy_config_defender_allowfullscanonmappednetworkdrives"
                        choiceSettingValue = @{
                            value = "device_vendor_msft_policy_config_defender_allowfullscanonmappednetworkdrives_1"
                            children = @()
                        }
                    }
                },
                @{
                    id = "2"
                    settingInstance = @{
                        "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                        settingDefinitionId = "device_vendor_msft_policy_config_defender_allowioavprotection"
                        choiceSettingValue = @{
                            value = "device_vendor_msft_policy_config_defender_allowioavprotection_1"
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

function New-DefenderAntivirusPolicy {
    Write-LogMessage -Message "Creating Defender Antivirus policy..." -Type Info
    
    try {
        $body = @{
            name = "Defender Antivirus Policy"
            description = "Microsoft Defender Antivirus comprehensive configuration"
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
                        choiceSettingValue = @{
                            value = "device_vendor_msft_policy_config_defender_puaprotection_1"
                            children = @()
                        }
                    }
                },
                @{
                    id = "1"
                    settingInstance = @{
                        "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                        settingDefinitionId = "device_vendor_msft_policy_config_defender_cloudblocklevel"
                        choiceSettingValue = @{
                            value = "device_vendor_msft_policy_config_defender_cloudblocklevel_2"
                            children = @()
                        }
                    }
                },
                @{
                    id = "2"
                    settingInstance = @{
                        "@odata.type" = "#microsoft.graph.deviceManagementConfigurationSimpleSettingInstance"
                        settingDefinitionId = "device_vendor_msft_policy_config_defender_cloudextendedtimeout"
                        simpleSettingValue = @{
                            "@odata.type" = "#microsoft.graph.deviceManagementConfigurationIntegerSettingValue"
                            value = 50
                        }
                    }
                }
            )
        }
        
        $result = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies" -Body $body
        Write-LogMessage -Message "Created Defender Antivirus policy" -Type Success
        return $result
    }
    catch {
        Write-LogMessage -Message "Failed to create Defender Antivirus policy - $($_.Exception.Message)" -Type Error
        return $null
    }
}

function New-BitLockerPolicy {
    Write-LogMessage -Message "Creating BitLocker policy..." -Type Info
    
    try {
        $body = @{
            name = "BitLocker Encryption"
            description = "BitLocker drive encryption configuration"
            platforms = "windows10"
            technologies = "mdm"
            templateReference = @{
                templateId = ""
                templateFamily = "none"
            }
            settings = @(
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

function New-LAPSPolicy {
    Write-LogMessage -Message "Creating LAPS policy..." -Type Info
    
    try {
        $body = @{
            name = "LAPS Configuration"
            description = "Local Administrator Password Solution"
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
                        choiceSettingValue = @{
                            value = "device_vendor_msft_laps_policies_backupdirectory_1"
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
                }
            )
        }
        
        $result = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies" -Body $body
        Write-LogMessage -Message "Created LAPS policy" -Type Success
        return $result
    }
    catch {
        Write-LogMessage -Message "Failed to create LAPS policy - $($_.Exception.Message)" -Type Error
        return $null
    }
}

function New-OneDrivePolicy {
    Write-LogMessage -Message "Creating OneDrive configuration policy..." -Type Info
    
    try {
        $body = @{
            name = "OneDrive Configuration"
            description = "OneDrive for Business configuration and Known Folder Move"
            platforms = "windows10"
            technologies = "mdm"
            templateReference = @{
                templateId = ""
                templateFamily = "none"
            }
            settings = @(
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

function New-EdgePolicies {
    Write-LogMessage -Message "Creating Microsoft Edge policies..." -Type Info
    
    try {
        # Get SharePoint root site URL
        $sharePointUrl = Get-SharePointRootSiteUrl
        
        $body = @{
            name = "Microsoft Edge Configuration"
            description = "Edge browser configuration with default pages"
            platforms = "windows10"
            technologies = "mdm"
            templateReference = @{
                templateId = ""
                templateFamily = "none"
            }
            settings = @(
                @{
                    id = "0"
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
                }
            )
        }
        
        $result = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies" -Body $body
        Write-LogMessage -Message "Created Microsoft Edge policy with SharePoint home page: $sharePointUrl" -Type Success
        return $result
    }
    catch {
        Write-LogMessage -Message "Failed to create Edge policies - $($_.Exception.Message)" -Type Error
        return $null
    }
}

function Get-SharePointRootSiteUrl {
    try {
        # First try to get from tenant state if available
        if ($script:TenantState -and $script:TenantState.DefaultDomain) {
            $domain = $script:TenantState.DefaultDomain
            $tenantName = $domain.Split('.')[0]
            $rootUrl = "https://$tenantName.sharepoint.com"
            return $rootUrl
        }
        
        # Fallback: use default
        Write-LogMessage -Message "Could not determine SharePoint URL, using default" -Type Warning
        return "https://www.office.com"
    }
    catch {
        Write-LogMessage -Message "Error determining SharePoint URL - $($_.Exception.Message)" -Type Warning
        return "https://www.office.com"
    }
}

function New-PowerOptionsPolicy {
    Write-LogMessage -Message "Creating Power Options policy..." -Type Info
    
    try {
        $body = @{
            name = "Power Options"
            description = "Power management settings for devices"
            platforms = "windows10"
            technologies = "mdm"
            templateReference = @{
                templateId = ""
                templateFamily = "none"
            }
            settings = @(
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

function New-AdminAccountPolicy {
    Write-LogMessage -Message "Creating Admin Account policy..." -Type Info
    
    try {
        $body = @{
            name = "Enable Built-in Administrator Account"
            description = "Enable and configure built-in administrator account for LAPS"
            platforms = "windows10"
            technologies = "mdm"
            templateReference = @{
                templateId = ""
                templateFamily = "none"
            }
            settings = @(
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
                }
            )
        }
        
        $result = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies" -Body $body
        Write-LogMessage -Message "Created Admin Account policy" -Type Success
        return $result
    }
    catch {
        Write-LogMessage -Message "Failed to create Admin Account policy - $($_.Exception.Message)" -Type Error
        return $null
    }
}

function New-FirewallPolicy {
    Write-LogMessage -Message "Creating Windows Firewall policy..." -Type Info
    
    try {
        $body = @{
            name = "Windows Firewall Configuration"
            description = "Default Windows Firewall settings for all network profiles"
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
                        choiceSettingValue = @{
                            value = "vendor_msft_firewall_mdmstore_domainprofile_enablefirewall_true"
                            children = @()
                        }
                    }
                }
            )
        }
        
        $result = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies" -Body $body
        Write-LogMessage -Message "Created Windows Firewall policy" -Type Success
        return $result
    }
    catch {
        Write-LogMessage -Message "Failed to create Windows Firewall policy - $($_.Exception.Message)" -Type Error
        return $null
    }
}

function New-TamperProtectionPolicy {
    Write-LogMessage -Message "Creating Tamper Protection policy..." -Type Info
    
    try {
        $body = @{
            name = "Tamper Protection"
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
                        choiceSettingValue = @{
                            value = "vendor_msft_defender_configuration_tamperprotection_options_0"
                            children = @()
                        }
                    }
                }
            )
        }
        
        $result = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies" -Body $body
        Write-LogMessage -Message "Created Tamper Protection policy" -Type Success
        return $result
    }
    catch {
        Write-LogMessage -Message "Failed to create Tamper Protection policy - $($_.Exception.Message)" -Type Error
        return $null
    }
}

function New-EDRPolicy {
    Write-LogMessage -Message "Creating EDR policy..." -Type Info
    
    try {
        $body = @{
            name = "EDR Configuration"
            description = "Endpoint Detection and Response configuration"
            platforms = "windows10"
            technologies = "mdm,microsoftSense"
            templateReference = @{
                templateId = "0385b795-0f2f-44ac-8602-9f65bf6adede_1"
                templateFamily = "endpointSecurityEndpointDetectionAndResponse"
                templateDisplayName = "Endpoint detection and response"
                templateDisplayVersion = "Version 1"
            }
            settings = @(
                @{
                    id = "0"
                    settingInstance = @{
                        "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                        settingDefinitionId = "device_vendor_msft_windowsadvancedthreatprotection_configurationtype"
                        choiceSettingValue = @{
                            value = "device_vendor_msft_windowsadvancedthreatprotection_configurationtype_autofromconnector"
                            children = @()
                        }
                    }
                }
            )
        }
        
        $result = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies" -Body $body
        Write-LogMessage -Message "Created EDR policy" -Type Success
        return $result
    }
    catch {
        Write-LogMessage -Message "Failed to create EDR policy - $($_.Exception.Message)" -Type Error
        return $null
    }
}

function New-OfficePolicies {
    Write-LogMessage -Message "Creating Office configuration policies..." -Type Info
    
    try {
        $body = @{
            name = "Office Updates Configuration"
            description = "Microsoft Office update settings"
            platforms = "windows10"
            technologies = "mdm"
            templateReference = @{
                templateId = ""
                templateFamily = "none"
            }
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
        return $result
    }
    catch {
        Write-LogMessage -Message "Failed to create Office policies - $($_.Exception.Message)" -Type Error
        return $null
    }
}

function New-OutlookPolicy {
    Write-LogMessage -Message "Creating Outlook configuration policy..." -Type Info
    
    try {
        $body = @{
            name = "Outlook Configuration"
            description = "Microsoft Outlook user experience settings"
            platforms = "windows10"
            technologies = "mdm"
            templateReference = @{
                templateId = ""
                templateFamily = "none"
            }
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
        return $result
    }
    catch {
        Write-LogMessage -Message "Failed to create Outlook policy - $($_.Exception.Message)" -Type Error
        return $null
    }
}

function New-UnenrollmentPolicy {
    Write-LogMessage -Message "Creating Device Unenrollment Prevention policy..." -Type Info
    
    try {
        $body = @{
            name = "Prevent Users From Unenrolling Devices"
            description = "Prevent users from manually unenrolling devices from Intune"
            platforms = "windows10"
            technologies = "mdm"
            templateReference = @{
                templateId = ""
                templateFamily = "none"
            }
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
        return $result
    }
    catch {
        Write-LogMessage -Message "Failed to create Unenrollment Prevention policy - $($_.Exception.Message)" -Type Error
        return $null
    }
}

# === Assignment Functions ===

function Assign-PolicyToGroups {
    param (
        [string]$PolicyId,
        [array]$GroupNames
    )
    
    try {
        $assignments = @()
        
        foreach ($groupName in $GroupNames) {
            if ($script:TenantState.CreatedGroups.ContainsKey($groupName)) {
                $groupId = $script:TenantState.CreatedGroups[$groupName]
                $assignments += @{
                    target = @{
                        "@odata.type" = "#microsoft.graph.groupAssignmentTarget"
                        groupId = $groupId
                    }
                }
            }
        }
        
        if ($assignments.Count -gt 0) {
            $body = @{
                assignments = $assignments
            }
            
            Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies/$PolicyId/assignments" -Body $body
            Write-LogMessage -Message "Assigned policy $PolicyId to $($assignments.Count) groups" -Type Success
        }
    }
    catch {
        Write-LogMessage -Message "Failed to assign policy $PolicyId - $($_.Exception.Message)" -Type Warning
    }
}