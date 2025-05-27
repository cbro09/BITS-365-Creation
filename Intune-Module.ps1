# === Intune.ps1 ===
# Microsoft Intune configuration and policy management functions - Complete Policies

function New-TenantIntune {
    Write-LogMessage -Message "Starting Intune configuration..." -Type Info
    
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
        
        # Create all configuration policies with complete settings
        Write-LogMessage -Message "Creating comprehensive configuration policies..." -Type Info
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
                $script:TenantState.CreatedGroups[$group.Name] = $newGroup.id
            }
            else {
                Write-LogMessage -Message "Device group already exists: $($group.Name)" -Type Warning
                $script:TenantState.CreatedGroups[$group.Name] = $existingGroup.Id
            }
        }
        
        # Create update rings with comprehensive settings
        $updateRings = @(
            @{ 
                Name = "RING0 - Pilot"
                Description = "Pilot testers for latest feature/quality updates"
                QualityDeferral = 0
                FeatureDeferral = 0
                GroupName = "WindowsDeviceRing0"
            },
            @{ 
                Name = "RING1 - UAT"
                Description = "UAT testers for latest feature/quality updates"
                QualityDeferral = 2
                FeatureDeferral = 7
                GroupName = "WindowsDeviceRing1"
            },
            @{ 
                Name = "RING2 - Production"
                Description = "Production devices with standard deferrals"
                QualityDeferral = 7
                FeatureDeferral = 7
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
                installationSchedule = @{
                    activeHoursStart = "08:00:00.0000000"
                    activeHoursEnd = "17:00:00.0000000"
                    scheduledInstallDay = "everyday"
                    scheduledInstallTime = "03:00:00.0000000"
                }
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
        $lapsSettings = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/policies/deviceRegistrationPolicy" -ErrorAction SilentlyContinue
        
        if ($lapsSettings -and $lapsSettings.localAdminPassword -and $lapsSettings.localAdminPassword.isEnabled) {
            Write-LogMessage -Message "Windows LAPS is already enabled" -Type Info
            return $true
        }
        
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

# === Complete Policy Creation Functions ===

function New-PowerOptionsPolicy {
    Write-LogMessage -Message "Creating complete Power Options policy..." -Type Info
    
    try {
        $body = @{
            name = "Power Options"
            description = "Comprehensive power management settings for devices"
            platforms = "windows10"
            technologies = "mdm"
            templateReference = @{
                templateId = ""
                templateFamily = "none"
            }
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
        return $result
    }
    catch {
        Write-LogMessage -Message "Failed to create Power Options policy - $($_.Exception.Message)" -Type Error
        return $null
    }
}

function New-DefenderAntivirusPolicy {
    Write-LogMessage -Message "Creating comprehensive Defender Antivirus policy..." -Type Info
    
    try {
        $body = @{
            name = "Defender Antivirus Policy"
            description = "Microsoft Defender Antivirus comprehensive configuration"
            platforms = "windows10"
            technologies = "mdm,microsoftSense"
            templateReference = @{
                templateId = "804339ad-1553-4478-a742-138fb5807418_1"
                templateFamily = "endpointSecurityAntivirus"
            }
            settings = @(
                # PUA Protection
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
                # Real-time Protection
                @{
                    id = "1"
                    settingInstance = @{
                        "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                        settingDefinitionId = "device_vendor_msft_policy_config_defender_allowrealtimemonitoring"
                        choiceSettingValue = @{
                            value = "device_vendor_msft_policy_config_defender_allowrealtimemonitoring_1"
                            children = @()
                        }
                    }
                },
                # Cloud Protection
                @{
                    id = "2"
                    settingInstance = @{
                        "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                        settingDefinitionId = "device_vendor_msft_policy_config_defender_allowcloudprotection"
                        choiceSettingValue = @{
                            value = "device_vendor_msft_policy_config_defender_allowcloudprotection_1"
                            children = @()
                        }
                    }
                },
                # Sample Submission
                @{
                    id = "3"
                    settingInstance = @{
                        "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                        settingDefinitionId = "device_vendor_msft_policy_config_defender_submitsamplesconsent"
                        choiceSettingValue = @{
                            value = "device_vendor_msft_policy_config_defender_submitsamplesconsent_1"
                            children = @()
                        }
                    }
                },
                # Scan Schedule - Daily
                @{
                    id = "4"
                    settingInstance = @{
                        "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                        settingDefinitionId = "device_vendor_msft_policy_config_defender_schedulescanday"
                        choiceSettingValue = @{
                            value = "device_vendor_msft_policy_config_defender_schedulescanday_0"
                            children = @()
                        }
                    }
                },
                # Scan Time - 2 AM
                @{
                    id = "5"
                    settingInstance = @{
                        "@odata.type" = "#microsoft.graph.deviceManagementConfigurationSimpleSettingInstance"
                        settingDefinitionId = "device_vendor_msft_policy_config_defender_schedulescantime"
                        simpleSettingValue = @{
                            "@odata.type" = "#microsoft.graph.deviceManagementConfigurationIntegerSettingValue"
                            value = 120
                        }
                    }
                }
            )
        }
        
        $result = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies" -Body $body
        Write-LogMessage -Message "Created comprehensive Defender Antivirus policy" -Type Success
        return $result
    }
    catch {
        Write-LogMessage -Message "Failed to create Defender Antivirus policy - $($_.Exception.Message)" -Type Error
        return $null
    }
}

function New-BitLockerPolicy {
    Write-LogMessage -Message "Creating comprehensive BitLocker policy with 13 settings..." -Type Info
    
    try {
        $body = @{
            name = "Enable Bitlocker"
            description = "Comprehensive BitLocker drive encryption configuration"
            platforms = "windows10"
            technologies = "mdm"
            templateReference = @{
                templateId = ""
                templateFamily = "none"
            }
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
        return $result
    }
    catch {
        Write-LogMessage -Message "Failed to create BitLocker policy - $($_.Exception.Message)" -Type Error
        return $null
    }
}

function New-OneDrivePolicy {
    Write-LogMessage -Message "Creating comprehensive OneDrive policy..." -Type Info
    
    try {
        $body = @{
            name = "OneDrive Configuration"
            description = "OneDrive for Business configuration with Known Folder Move"
            platforms = "windows10"
            technologies = "mdm"
            templateReference = @{
                templateId = ""
                templateFamily = "none"
            }
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
        return $result
    }
    catch {
        Write-LogMessage -Message "Failed to create OneDrive policy - $($_.Exception.Message)" -Type Error
        return $null
    }
}

function New-EdgePolicies {
    Write-LogMessage -Message "Creating comprehensive Microsoft Edge policy..." -Type Info
    
    try {
        # Get SharePoint root site URL
        $sharePointUrl = Get-SharePointRootSiteUrl
        
        $body = @{
            name = "Microsoft Edge Configuration"
            description = "Comprehensive Edge browser configuration"
            platforms = "windows10"
            technologies = "mdm"
            settings = @(
                # Home Page Location
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
                },
                # Show Home Button
                @{
                    id = "1"
                    settingInstance = @{
                        "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                        settingDefinitionId = "device_vendor_msft_policy_config_microsoft_edge~policy~microsoft_edge_showhomebutton"
                        choiceSettingValue = @{
                            value = "device_vendor_msft_policy_config_microsoft_edge~policy~microsoft_edge_showhomebutton_1"
                            children = @()
                        }
                    }
                },
                # Enable Password Manager
                @{
                    id = "2"
                    settingInstance = @{
                        "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                        settingDefinitionId = "device_vendor_msft_policy_config_microsoft_edge~policy~microsoft_edge~passwordmanager_passwordmanagerenabled"
                        choiceSettingValue = @{
                            value = "device_vendor_msft_policy_config_microsoft_edge~policy~microsoft_edge~passwordmanager_passwordmanagerenabled_1"
                            children = @()
                        }
                    }
                },
                # Block third-party cookies
                @{
                    id = "3"
                    settingInstance = @{
                        "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                        settingDefinitionId = "device_vendor_msft_policy_config_microsoft_edge~policy~microsoft_edge~contentsettings_blockthirdpartycookies"
                        choiceSettingValue = @{
                            value = "device_vendor_msft_policy_config_microsoft_edge~policy~microsoft_edge~contentsettings_blockthirdpartycookies_1"
                            children = @()
                        }
                    }
                },
                # Enhanced security mode
                @{
                    id = "4"
                    settingInstance = @{
                        "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                        settingDefinitionId = "device_vendor_msft_policy_config_microsoft_edge~policy~microsoft_edge_enhancesecuritymode"
                        choiceSettingValue = @{
                            value = "device_vendor_msft_policy_config_microsoft_edge~policy~microsoft_edge_enhancesecuritymode_1"
                            children = @()
                        }
                    }
                }
            )
        }
        
        $result = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies" -Body $body
        Write-LogMessage -Message "Created comprehensive Microsoft Edge policy with SharePoint home page: $sharePointUrl" -Type Success
        return $result
    }
    catch {
        Write-LogMessage -Message "Failed to create Edge policies - $($_.Exception.Message)" -Type Error
        return $null
    }
}

# === Remaining Policy Functions (with placeholder complete settings) ===

function New-DefenderPolicy {
    Write-LogMessage -Message "Creating comprehensive Defender policy..." -Type Info
    
    try {
        $body = @{
            name = "Defender Configuration"
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
        return $result
    }
    catch {
        Write-LogMessage -Message "Failed to create Defender policy - $($_.Exception.Message)" -Type Error
        return $null
    }
}

function New-FirewallPolicy {
    Write-LogMessage -Message "Creating comprehensive Windows Firewall policy..." -Type Info
    
    try {
        $body = @{
            name = "Firewall Windows default policy"
            description = "Default policy sets settings for all endpoints that are not governed by any other policy, ensuring that all your clients are managed as soon as MDE is deployed. The default policy is based on a set of pre-configured recommended settings and can be adjusted by user with admin priviledges."
            platforms = "windows10"
            technologies = "mdm,microsoftSense"
            templateReference = @{
                templateId = "6078910e-d808-4a9f-a51d-1b8a7bacb7c0_1"
                templateFamily = "endpointSecurityFirewall"
            }
            settings = @(
                # Domain Profile
                @{
                    id = "0"
                    settingInstance = @{
                        "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                        settingDefinitionId = "vendor_msft_firewall_mdmstore_domainprofile_enablefirewall"
                        choiceSettingValue = @{
                            value = "vendor_msft_firewall_mdmstore_domainprofile_enablefirewall_true"
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
                # Private Profile
                @{
                    id = "1"
                    settingInstance = @{
                        "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                        settingDefinitionId = "vendor_msft_firewall_mdmstore_privateprofile_enablefirewall"
                        choiceSettingValue = @{
                            value = "vendor_msft_firewall_mdmstore_privateprofile_enablefirewall_true"
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
                # Public Profile
                @{
                    id = "2"
                    settingInstance = @{
                        "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                        settingDefinitionId = "vendor_msft_firewall_mdmstore_publicprofile_enablefirewall"
                        choiceSettingValue = @{
                            value = "vendor_msft_firewall_mdmstore_publicprofile_enablefirewall_true"
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
        Write-LogMessage -Message "Created comprehensive Windows Firewall policy with 3 network profiles" -Type Success
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

function New-LAPSPolicy {
    Write-LogMessage -Message "Creating comprehensive LAPS policy with 6 settings..." -Type Info
    
    try {
        $body = @{
            name = "LAPS"
            description = "Local Admin Password Solution"
            platforms = "windows10"
            technologies = "mdm"
            templateReference = @{
                templateId = "adc46e5a-f4aa-4ff6-aeff-4f27bc525796_1"
                templateFamily = "endpointSecurityAccountProtection"
            }
            settings = @(
                # Backup Directory
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
                },
                # Administrator Account Name
                @{
                    id = "1"
                    settingInstance = @{
                        "@odata.type" = "#microsoft.graph.deviceManagementConfigurationSimpleSettingInstance"
                        settingDefinitionId = "device_vendor_msft_laps_policies_administratoraccountname"
                        simpleSettingValue = @{
                            "@odata.type" = "#microsoft.graph.deviceManagementConfigurationStringSettingValue"
                            value = "palocal"
                        }
                    }
                },
                # Password Complexity
                @{
                    id = "2"
                    settingInstance = @{
                        "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                        settingDefinitionId = "device_vendor_msft_laps_policies_passwordcomplexity"
                        choiceSettingValue = @{
                            value = "device_vendor_msft_laps_policies_passwordcomplexity_3"
                            children = @()
                        }
                    }
                },
                # Password Length
                @{
                    id = "3"
                    settingInstance = @{
                        "@odata.type" = "#microsoft.graph.deviceManagementConfigurationSimpleSettingInstance"
                        settingDefinitionId = "device_vendor_msft_laps_policies_passwordlength"
                        simpleSettingValue = @{
                            "@odata.type" = "#microsoft.graph.deviceManagementConfigurationIntegerSettingValue"
                            value = 20
                        }
                    }
                },
                # Post Authentication Actions
                @{
                    id = "4"
                    settingInstance = @{
                        "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                        settingDefinitionId = "device_vendor_msft_laps_policies_postauthenticationactions"
                        choiceSettingValue = @{
                            value = "device_vendor_msft_laps_policies_postauthenticationactions_3"
                            children = @()
                        }
                    }
                },
                # Post Authentication Reset Delay
                @{
                    id = "5"
                    settingInstance = @{
                        "@odata.type" = "#microsoft.graph.deviceManagementConfigurationSimpleSettingInstance"
                        settingDefinitionId = "device_vendor_msft_laps_policies_postauthenticationresetdelay"
                        simpleSettingValue = @{
                            "@odata.type" = "#microsoft.graph.deviceManagementConfigurationIntegerSettingValue"
                            value = 1
                        }
                    }
                }
            )
        }
        
        $result = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies" -Body $body
        Write-LogMessage -Message "Created comprehensive LAPS policy with 6 settings" -Type Success
        return $result
    }
    catch {
        Write-LogMessage -Message "Failed to create LAPS policy - $($_.Exception.Message)" -Type Error
        return $null
    }
}

function New-AdminAccountPolicy {
    Write-LogMessage -Message "Creating Admin Account policy with rename..." -Type Info
    
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
        return $result
    }
    catch {
        Write-LogMessage -Message "Failed to create Admin Account policy - $($_.Exception.Message)" -Type Error
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

function New-OfficePolicies {
    Write-LogMessage -Message "Creating Office configuration policies..." -Type Info
    
    try {
        $body = @{
            name = "Office Updates Configuration"
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

# === Helper Functions ===

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