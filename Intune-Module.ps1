# === Intune.ps1 ===
# Microsoft Intune configuration and policy management functions - Complete Implementation

function New-TenantIntune {
    Write-LogMessage -Message "Starting comprehensive Intune configuration..." -Type Info
    
    try {
        # Store core functions to prevent them being cleared during module reloading
        $writeLogFunction = ${function:Write-LogMessage}
        $testNotEmptyFunction = ${function:Test-NotEmpty}
        $showProgressFunction = ${function:Show-Progress}
        
        # Remove ALL Graph modules first to avoid conflicts
        Write-LogMessage -Message "Clearing all Graph modules to prevent conflicts..." -Type Info
        Get-Module Microsoft.Graph* | Remove-Module -Force -ErrorAction SilentlyContinue
        
        # Restore core functions after module cleanup
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
        
        # Load only the essential modules for Intune operations
        $intuneModules = @(
            'Microsoft.Graph.DeviceManagement',
            'Microsoft.Graph.Groups', 
            'Microsoft.Graph.Identity.DirectoryManagement'
        )
        
        Write-LogMessage -Message "Loading essential Intune modules..." -Type Info
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
        
        # Connect with required scopes for Intune management
        $intuneScopes = @(
            "DeviceManagementConfiguration.ReadWrite.All",
            "DeviceManagementManagedDevices.ReadWrite.All", 
            "DeviceManagementApps.ReadWrite.All",
            "Group.ReadWrite.All",
            "Directory.ReadWrite.All"
        )
        
        Write-LogMessage -Message "Connecting to Microsoft Graph with Intune permissions..." -Type Info
        Connect-MgGraph -Scopes $intuneScopes -ErrorAction Stop | Out-Null
        
        $context = Get-MgContext
        Write-LogMessage -Message "Connected to Microsoft Graph as $($context.Account)" -Type Success
        Write-LogMessage -Message "Active scopes: $($context.Scopes -join ', ')" -Type Info -LogOnly
        
        # Verify all required scopes are present
        $missingScopes = @()
        foreach ($scope in $intuneScopes) {
            if ($context.Scopes -notcontains $scope) {
                $missingScopes += $scope
            }
        }
        
        if ($missingScopes.Count -gt 0) {
            Write-LogMessage -Message "Missing required scopes: $($missingScopes -join ', ')" -Type Error
            return $false
        }
        
        Write-LogMessage -Message "Graph connection verified with all required permissions" -Type Success
        
        # Optional: Clean up duplicate policies from previous runs
        Write-Host ""
        $cleanup = Read-Host "Would you like to check for and remove duplicate policies first? (Y/N)"
        if ($cleanup -eq 'Y' -or $cleanup -eq 'y') {
            $duplicatesRemoved = Remove-DuplicatePolicies
            Write-LogMessage -Message "Cleanup completed: $duplicatesRemoved duplicates removed" -Type Info
        }
        
        # 1. Create WindowsAutoPilot dynamic group
        $autopilotGroup = New-WindowsAutoPilotGroup
        if (-not $autopilotGroup) {
            Write-LogMessage -Message "Failed to create WindowsAutoPilot group" -Type Warning
        }
        
        # 2. Create Windows Update rings and device groups
        $updateRings = New-WindowsUpdateRings
        if (-not $updateRings) {
            Write-LogMessage -Message "Failed to create Windows Update rings" -Type Warning
        }
        
        # 3. Enable LAPS prerequisite
        $lapsEnabled = Enable-WindowsLAPS
        if (-not $lapsEnabled) {
            Write-LogMessage -Message "LAPS enablement failed - LAPS policies may not work correctly" -Type Warning
        }
        else {
            Write-LogMessage -Message "Windows LAPS successfully enabled" -Type Success
        }
        
        # 4. Create comprehensive configuration policies
        Write-LogMessage -Message "Creating comprehensive Intune configuration policies..." -Type Info
        $policies = @()
        $policyResults = @{}
        
        # Security policies
        Write-LogMessage -Message "Creating security policies..." -Type Info
        $policies += New-DefenderAntivirusPolicy
        $policies += New-DefenderPolicy
        $policies += New-FirewallPolicy
        $policies += New-TamperProtectionPolicy
        
        # Show EDR setup instructions (requires manual configuration)
        Show-EDREnablementNote
        
        # Encryption policies
        Write-LogMessage -Message "Creating encryption policies..." -Type Info
        $policies += New-BitLockerPolicy
        
        # Authentication and access policies
        Write-LogMessage -Message "Creating authentication policies..." -Type Info
        $policies += New-LAPSPolicy
        $policies += New-AdminAccountPolicy
        
        # Device configuration policies
        Write-LogMessage -Message "Creating device configuration policies..." -Type Info
        $policies += New-OneDrivePolicy
        $policies += New-PowerOptionsPolicy
        $policies += New-UnenrollmentPolicy
        
        # Application configuration policies
        Write-LogMessage -Message "Creating application policies..." -Type Info
        $policies += New-EdgePolicies
        $policies += New-EdgeUpdatePolicy
        $policies += New-OfficePolicies
        $policies += New-OutlookPolicy
        $policies += New-DisableUACPolicy
        
        # Count successful policy creations
        $successfulPolicies = ($policies | Where-Object { $_ -and $_.id -and $_.id -ne "existing" -and $_.id -ne "manual_setup_required" }).Count
        $existingPolicies = ($policies | Where-Object { $_ -and $_.id -eq "existing" }).Count
        $totalPolicies = ($policies | Where-Object { $_ }).Count
        
        Write-LogMessage -Message "Policy creation summary: $successfulPolicies created, $existingPolicies already existed, $totalPolicies total" -Type Info
        
        # 5. Assign policies to device groups
        Write-LogMessage -Message "Assigning policies to device groups..." -Type Info
        $deviceGroups = @("WindowsDeviceRing0", "WindowsDeviceRing1", "WindowsDeviceRing2")
        $assignmentCount = 0
        
        foreach ($policy in $policies) {
            if ($policy -and $policy.id -and $policy.id -ne "existing" -and $policy.id -ne "manual_setup_required") {
                $assigned = Assign-PolicyToGroups -PolicyId $policy.id -GroupNames $deviceGroups
                if ($assigned) {
                    $assignmentCount++
                }
            }
        }
        
        Write-LogMessage -Message "Successfully assigned $assignmentCount policies to device groups" -Type Success
        Write-LogMessage -Message "Intune configuration completed successfully!" -Type Success
        
        # Final summary
        Write-LogMessage -Message "=== INTUNE CONFIGURATION SUMMARY ===" -Type Info
        Write-LogMessage -Message "✓ WindowsAutoPilot group: $(if($autopilotGroup) { 'Created/Verified' } else { 'Failed' })" -Type Info
        Write-LogMessage -Message "✓ Windows Update rings: $(if($updateRings) { 'Created/Verified' } else { 'Failed' })" -Type Info
        Write-LogMessage -Message "✓ Windows LAPS: $(if($lapsEnabled) { 'Enabled' } else { 'Failed' })" -Type Info
        Write-LogMessage -Message "✓ Configuration policies: $successfulPolicies created, $existingPolicies existing" -Type Info
        Write-LogMessage -Message "✓ Policy assignments: $assignmentCount policies assigned to device groups" -Type Info
        Write-LogMessage -Message "! EDR Policy: Requires manual setup (see instructions above)" -Type Warning
        Write-LogMessage -Message "=================================" -Type Info
        
        return $true
    }
    catch {
        Write-LogMessage -Message "Error in Intune configuration - $($_.Exception.Message)" -Type Error
        Write-LogMessage -Message "Stack trace: $($_.Exception.StackTrace)" -Type Error -LogOnly
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
            $script:TenantState.CreatedGroups["WindowsAutoPilot"] = $existingGroup.Id
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
        $script:TenantState.CreatedGroups["WindowsAutoPilot"] = $result.id
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
        
        # Create Windows Update rings with comprehensive settings
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
            # Check if update ring already exists
            try {
                $existingRings = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/deviceManagement/windowsUpdateForBusinessRings" -ErrorAction Stop
                $existingRing = $existingRings.value | Where-Object { $_.displayName -eq $ring.Name }
                
                if ($existingRing) {
                    Write-LogMessage -Message "Update ring already exists: $($ring.Name)" -Type Warning
                    continue
                }
            }
            catch {
                Write-LogMessage -Message "Error checking existing update rings: $($_.Exception.Message)" -Type Warning
            }
            
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
                    
                    try {
                        Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/windowsUpdateForBusinessRings/$($result.id)/assignments" -Body $assignmentBody
                        Write-LogMessage -Message "Assigned update ring $($ring.Name) to group $($ring.GroupName)" -Type Success
                    }
                    catch {
                        Write-LogMessage -Message "Failed to assign update ring $($ring.Name) to group: $($_.Exception.Message)" -Type Warning
                    }
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

# === Comprehensive Policy Creation Functions ===

function New-DefenderAntivirusPolicy {
    Write-LogMessage -Message "Creating comprehensive Defender Antivirus policy with 27 settings..." -Type Info
    
    $policyName = "NGP Windows default policy"
    if (Test-PolicyExists -PolicyName $policyName) {
        Write-LogMessage -Message "Policy '$policyName' already exists, skipping creation" -Type Warning
        return @{ name = $policyName; id = "existing" }
    }
    
    try {
        $body = @{
            name = $policyName
            description = "Default policy sets settings for all endpoints that are not governed by any other policy, ensuring that all your clients are managed as soon as MDE is deployed. The default policy is based on a set of pre-configured recommended settings and can be adjusted by user with admin privileges."
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
                },
                @{
                    id = "3"
                    settingInstance = @{
                        "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                        settingDefinitionId = "device_vendor_msft_policy_config_defender_enablenetworkprotection"
                        choiceSettingValue = @{
                            value = "device_vendor_msft_policy_config_defender_enablenetworkprotection_1"
                            children = @()
                        }
                    }
                },
                @{
                    id = "4"
                    settingInstance = @{
                        "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                        settingDefinitionId = "device_vendor_msft_policy_config_defender_allowrealtimemonitoring"
                        choiceSettingValue = @{
                            value = "device_vendor_msft_policy_config_defender_allowrealtimemonitoring_1"
                            children = @()
                        }
                    }
                },
                @{
                    id = "5"
                    settingInstance = @{
                        "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                        settingDefinitionId = "device_vendor_msft_policy_config_defender_schedulescanday"
                        choiceSettingValue = @{
                            value = "device_vendor_msft_policy_config_defender_schedulescanday_2"
                            children = @()
                        }
                    }
                },
                @{
                    id = "6"
                    settingInstance = @{
                        "@odata.type" = "#microsoft.graph.deviceManagementConfigurationSimpleSettingInstance"
                        settingDefinitionId = "device_vendor_msft_policy_config_defender_schedulequickscantime"
                        simpleSettingValue = @{
                            "@odata.type" = "#microsoft.graph.deviceManagementConfigurationIntegerSettingValue"
                            value = 720
                        }
                    }
                },
                @{
                    id = "7"
                    settingInstance = @{
                        "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                        settingDefinitionId = "device_vendor_msft_policy_config_defender_scanparameter"
                        choiceSettingValue = @{
                            value = "device_vendor_msft_policy_config_defender_scanparameter_1"
                            children = @()
                        }
                    }
                },
                @{
                    id = "8"
                    settingInstance = @{
                        "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                        settingDefinitionId = "device_vendor_msft_policy_config_defender_allowarchivescanning"
                        choiceSettingValue = @{
                            value = "device_vendor_msft_policy_config_defender_allowarchivescanning_1"
                            children = @()
                        }
                    }
                },
                @{
                    id = "9"
                    settingInstance = @{
                        "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                        settingDefinitionId = "device_vendor_msft_policy_config_defender_enablelowcpupriority"
                        choiceSettingValue = @{
                            value = "device_vendor_msft_policy_config_defender_enablelowcpupriority_0"
                            children = @()
                        }
                    }
                },
                @{
                    id = "10"
                    settingInstance = @{
                        "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                        settingDefinitionId = "device_vendor_msft_policy_config_defender_disablecatchupfullscan"
                        choiceSettingValue = @{
                            value = "device_vendor_msft_policy_config_defender_disablecatchupfullscan_0"
                            children = @()
                        }
                    }
                },
                @{
                    id = "11"
                    settingInstance = @{
                        "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                        settingDefinitionId = "device_vendor_msft_policy_config_defender_disablecatchupquickscan"
                        choiceSettingValue = @{
                            value = "device_vendor_msft_policy_config_defender_disablecatchupquickscan_0"
                            children = @()
                        }
                    }
                },
                @{
                    id = "12"
                    settingInstance = @{
                        "@odata.type" = "#microsoft.graph.deviceManagementConfigurationSimpleSettingInstance"
                        settingDefinitionId = "device_vendor_msft_policy_config_defender_avgcpuloadfactor"
                        simpleSettingValue = @{
                            "@odata.type" = "#microsoft.graph.deviceManagementConfigurationIntegerSettingValue"
                            value = 50
                        }
                    }
                },
                @{
                    id = "13"
                    settingInstance = @{
                        "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                        settingDefinitionId = "device_vendor_msft_policy_config_defender_allowuseruiaccess"
                        choiceSettingValue = @{
                            value = "device_vendor_msft_policy_config_defender_allowuseruiaccess_1"
                            children = @()
                        }
                    }
                },
                @{
                    id = "14"
                    settingInstance = @{
                        "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                        settingDefinitionId = "device_vendor_msft_policy_config_defender_allowcloudprotection"
                        choiceSettingValue = @{
                            value = "device_vendor_msft_policy_config_defender_allowcloudprotection_1"
                            children = @()
                        }
                    }
                },
                @{
                    id = "15"
                    settingInstance = @{
                        "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                        settingDefinitionId = "device_vendor_msft_policy_config_defender_realtimescandirection"
                        choiceSettingValue = @{
                            value = "device_vendor_msft_policy_config_defender_realtimescandirection_0"
                            children = @()
                        }
                    }
                },
                @{
                    id = "16"
                    settingInstance = @{
                        "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                        settingDefinitionId = "device_vendor_msft_policy_config_defender_allowbehaviormonitoring"
                        choiceSettingValue = @{
                            value = "device_vendor_msft_policy_config_defender_allowbehaviormonitoring_1"
                            children = @()
                        }
                    }
                },
                @{
                    id = "17"
                    settingInstance = @{
                        "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                        settingDefinitionId = "device_vendor_msft_policy_config_defender_allowioavprotection"
                        choiceSettingValue = @{
                            value = "device_vendor_msft_policy_config_defender_allowioavprotection_1"
                            children = @()
                        }
                    }
                },
                @{
                    id = "18"
                    settingInstance = @{
                        "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                        settingDefinitionId = "device_vendor_msft_policy_config_defender_allowscriptscanning"
                        choiceSettingValue = @{
                            value = "device_vendor_msft_policy_config_defender_allowscriptscanning_1"
                            children = @()
                        }
                    }
                },
                @{
                    id = "19"
                    settingInstance = @{
                        "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                        settingDefinitionId = "device_vendor_msft_policy_config_defender_allowscanningnetworkfiles"
                        choiceSettingValue = @{
                            value = "device_vendor_msft_policy_config_defender_allowscanningnetworkfiles_1"
                            children = @()
                        }
                    }
                },
                @{
                    id = "20"
                    settingInstance = @{
                        "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                        settingDefinitionId = "device_vendor_msft_policy_config_defender_allowemailscanning"
                        choiceSettingValue = @{
                            value = "device_vendor_msft_policy_config_defender_allowemailscanning_0"
                            children = @()
                        }
                    }
                },
                @{
                    id = "21"
                    settingInstance = @{
                        "@odata.type" = "#microsoft.graph.deviceManagementConfigurationSimpleSettingInstance"
                        settingDefinitionId = "device_vendor_msft_policy_config_defender_daystoretaincleanedmalware"
                        simpleSettingValue = @{
                            "@odata.type" = "#microsoft.graph.deviceManagementConfigurationIntegerSettingValue"
                            value = 0
                        }
                    }
                },
                @{
                    id = "22"
                    settingInstance = @{
                        "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                        settingDefinitionId = "device_vendor_msft_policy_config_defender_submitsamplesconsent"
                        choiceSettingValue = @{
                            value = "device_vendor_msft_policy_config_defender_submitsamplesconsent_1"
                            children = @()
                        }
                    }
                },
                @{
                    id = "23"
                    settingInstance = @{
                        "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                        settingDefinitionId = "device_vendor_msft_policy_config_defender_allowfullscanonmappednetworkdrives"
                        choiceSettingValue = @{
                            value = "device_vendor_msft_policy_config_defender_allowfullscanonmappednetworkdrives_0"
                            children = @()
                        }
                    }
                },
                @{
                    id = "24"
                    settingInstance = @{
                        "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                        settingDefinitionId = "device_vendor_msft_policy_config_defender_allowfullscanremovabledrivescanning"
                        choiceSettingValue = @{
                            value = "device_vendor_msft_policy_config_defender_allowfullscanremovabledrivescanning_1"
                            children = @()
                        }
                    }
                },
                @{
                    id = "25"
                    settingInstance = @{
                        "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                        settingDefinitionId = "device_vendor_msft_policy_config_defender_checkforsignaturesbeforerunningscan"
                        choiceSettingValue = @{
                            value = "device_vendor_msft_policy_config_defender_checkforsignaturesbeforerunningscan_1"
                            children = @()
                        }
                    }
                },
                @{
                    id = "26"
                    settingInstance = @{
                        "@odata.type" = "#microsoft.graph.deviceManagementConfigurationSimpleSettingInstance"
                        settingDefinitionId = "device_vendor_msft_policy_config_defender_signatureupdateinterval"
                        simpleSettingValue = @{
                            "@odata.type" = "#microsoft.graph.deviceManagementConfigurationIntegerSettingValue"
                            value = 4
                        }
                    }
                }
            )
        }
        
        $result = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies" -Body $body
        Write-LogMessage -Message "Created $policyName with 27 comprehensive settings" -Type Success
        return $result
    }
    catch {
        Write-LogMessage -Message "Failed to create $policyName - $($_.Exception.Message)" -Type Error
        return $null
    }
}

function New-BitLockerPolicy {
    Write-LogMessage -Message "Creating comprehensive BitLocker policy with 13 settings..." -Type Info
    
    $policyName = "Enable Bitlocker"
    if (Test-PolicyExists -PolicyName $policyName) {
        Write-LogMessage -Message "Policy '$policyName' already exists, skipping creation" -Type Warning
        return @{ name = $policyName; id = "existing" }
    }
    
    try {
        $body = @{
            name = $policyName
            description = "Comprehensive BitLocker drive encryption configuration with recovery options"
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
                # Encryption method by drive type (XTS-AES 256)
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
                # System drives require startup authentication (TPM only)
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
                # System drives recovery options
                @{
                    id = "6"
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
                                    settingDefinitionId = "device_vendor_msft_bitlocker_systemdrivesrecoveryoptions_osactivedirectorybackup_name"
                                    choiceSettingValue = @{
                                        value = "device_vendor_msft_bitlocker_systemdrivesrecoveryoptions_osactivedirectorybackup_name_1"
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
        Write-LogMessage -Message "Created $policyName with comprehensive encryption settings" -Type Success
        return $result
    }
    catch {
        Write-LogMessage -Message "Failed to create $policyName - $($_.Exception.Message)" -Type Error
        return $null
    }
}

function New-OneDrivePolicy {
    Write-LogMessage -Message "Creating comprehensive OneDrive policy..." -Type Info
    
    $policyName = "OneDrive Configuration"
    if (Test-PolicyExists -PolicyName $policyName) {
        Write-LogMessage -Message "Policy '$policyName' already exists, skipping creation" -Type Warning
        return @{ name = $policyName; id = "existing" }
    }
    
    try {
        $body = @{
            name = $policyName
            description = "OneDrive for Business configuration with Known Folder Move and sync settings"
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
                        settingDefinitionId = "user_vendor_msft_policy_config_onedrivengscv2~policy~onedrivengsc_disablepauseonnetwork"
                        choiceSettingValue = @{
                            value = "user_vendor_msft_policy_config_onedrivengscv2~policy~onedrivengsc_disablepauseonnetwork_1"
                            children = @()
                        }
                    }
                },
                # Block opt-out from Known Folder Move
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
                # Silent Account Config
                @{
                    id = "3"
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
                    id = "4"
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
        Write-LogMessage -Message "Created comprehensive OneDrive policy with Known Folder Move" -Type Success
        return $result
    }
    catch {
        Write-LogMessage -Message "Failed to create OneDrive policy - $($_.Exception.Message)" -Type Error
        return $null
    }
}

function New-EdgePolicies {
    Write-LogMessage -Message "Creating Edge policy with SharePoint homepage..." -Type Info
    
    $policyName = "Default Web Pages"
    if (Test-PolicyExists -PolicyName $policyName) {
        Write-LogMessage -Message "Policy '$policyName' already exists, skipping creation" -Type Warning
        return @{ name = $policyName; id = "existing" }
    }
    
    try {
        # Get SharePoint root site URL from tenant
        $sharePointUrl = Get-SharePointRootSiteUrl
        
        $body = @{
            name = $policyName
            description = "Setting SharePoint home page as default startup page for Microsoft Edge"
            platforms = "windows10"
            technologies = "mdm"
            templateReference = @{
                templateId = ""
                templateFamily = "none"
            }
            settings = @(
                # Restore on startup
                @{
                    id = "0"
                    settingInstance = @{
                        "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                        settingDefinitionId = "device_vendor_msft_policy_config_microsoft_edgev77.3~policy~microsoft_edge~startup_restoreonstartup"
                        choiceSettingValue = @{
                            value = "device_vendor_msft_policy_config_microsoft_edgev77.3~policy~microsoft_edge~startup_restoreonstartup_1"
                            children = @(
                                @{
                                    "@odata.type" = "#microsoft.graph.deviceManagementConfigurationChoiceSettingInstance"
                                    settingDefinitionId = "device_vendor_msft_policy_config_microsoft_edgev77.3~policy~microsoft_edge~startup_restoreonstartup_restoreonstartup"
                                    choiceSettingValue = @{
                                        value = "device_vendor_msft_policy_config_microsoft_edgev77.3~policy~microsoft_edge~startup_restoreonstartup_restoreonstartup_5"
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
        return $result
    }
    catch {
        Write-LogMessage -Message "Failed to create Edge policies - $($_.Exception.Message)" -Type Error
        return $null
    }
}

function New-LAPSPolicy {
    Write-LogMessage -Message "Creating LAPS policy with tenant-specific admin name..." -Type Info
    
    $policyName = "LAPS"
    if (Test-PolicyExists -PolicyName $policyName) {
        Write-LogMessage -Message "Policy '$policyName' already exists, skipping creation" -Type Warning
        return @{ name = $policyName; id = "existing" }
    }
    
    try {
        # Generate admin account name based on tenant
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
            description = "Local Admin Password Solution with automatic password rotation"
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
                },
                @{
                    id = "1"
                    settingInstance = @{
                        "@odata.type" = "#microsoft.graph.deviceManagementConfigurationSimpleSettingInstance"
                        settingDefinitionId = "device_vendor_msft_laps_policies_administratoraccountname"
                        simpleSettingValue = @{
                            "@odata.type" = "#microsoft.graph.deviceManagementConfigurationStringSettingValue"
                            value = $adminAccountName
                        }
                    }
                },
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
        Write-LogMessage -Message "Created LAPS policy with admin account: $adminAccountName" -Type Success
        return $result
    }
    catch {
        Write-LogMessage -Message "Failed to create LAPS policy - $($_.Exception.Message)" -Type Error
        return $null
    }
}

# === Additional Policy Functions ===

function New-PowerOptionsPolicy {
    Write-LogMessage -Message "Creating comprehensive Power Options policy..." -Type Info
    
    $policyName = "Power Options"
    if (Test-PolicyExists -PolicyName $policyName) {
        Write-LogMessage -Message "Policy '$policyName' already exists, skipping creation" -Type Warning
        return @{ name = $policyName; id = "existing" }
    }
    
    try {
        $body = @{
            name = $policyName
            description = "Comprehensive power management settings for optimal device performance"
            platforms = "windows10"
            technologies = "mdm"
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
                },
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
                @{
                    id = "2"
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

function New-DefenderPolicy {
    Write-LogMessage -Message "Creating Defender configuration policy..." -Type Info
    
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
        Write-LogMessage -Message "Created Defender configuration policy" -Type Success
        return $result
    }
    catch {
        Write-LogMessage -Message "Failed to create Defender policy - $($_.Exception.Message)" -Type Error
        return $null
    }
}

function New-FirewallPolicy {
    Write-LogMessage -Message "Creating Windows Firewall policy..." -Type Info
    
    $policyName = "Firewall Windows default policy"
    if (Test-PolicyExists -PolicyName $policyName) {
        Write-LogMessage -Message "Policy '$policyName' already exists, skipping creation" -Type Warning
        return @{ name = $policyName; id = "existing" }
    }
    
    try {
        $body = @{
            name = $policyName
            description = "Windows Firewall configuration for all network profiles"
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
        Write-LogMessage -Message "Created Windows Firewall policy for all network profiles" -Type Success
        return $result
    }
    catch {
        Write-LogMessage -Message "Failed to create Windows Firewall policy - $($_.Exception.Message)" -Type Error
        return $null
    }
}

function New-TamperProtectionPolicy {
    Write-LogMessage -Message "Creating Tamper Protection policy..." -Type Info
    
    $policyName = "Tamper Protection"
    if (Test-PolicyExists -PolicyName $policyName) {
        Write-LogMessage -Message "Policy '$policyName' already exists, skipping creation" -Type Warning
        return @{ name = $policyName; id = "existing" }
    }
    
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

function New-AdminAccountPolicy {
    Write-LogMessage -Message "Creating Admin Account policy..." -Type Info
    
    $policyName = "Enable Built-in Administrator Account"
    if (Test-PolicyExists -PolicyName $policyName) {
        Write-LogMessage -Message "Policy '$policyName' already exists, skipping creation" -Type Warning
        return @{ name = $policyName; id = "existing" }
    }
    
    try {
        $body = @{
            name = $policyName
            description = "Enable and configure built-in administrator account for LAPS management"
            platforms = "windows10"
            technologies = "mdm"
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
                },
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
        Write-LogMessage -Message "Created Admin Account policy with rename to 'localadmin'" -Type Success
        return $result
    }
    catch {
        Write-LogMessage -Message "Failed to create Admin Account policy - $($_.Exception.Message)" -Type Error
        return $null
    }
}

function New-UnenrollmentPolicy {
    Write-LogMessage -Message "Creating Device Unenrollment Prevention policy..." -Type Info
    
    $policyName = "Prevent Users From Unenrolling Devices"
    if (Test-PolicyExists -PolicyName $policyName) {
        Write-LogMessage -Message "Policy '$policyName' already exists, skipping creation" -Type Warning
        return @{ name = $policyName; id = "existing" }
    }
    
    try {
        $body = @{
            name = $policyName
            description = "Prevent users from manually unenrolling devices from Intune management"
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

function New-EdgeUpdatePolicy {
    Write-LogMessage -Message "Creating Edge Update policy..." -Type Info
    
    $policyName = "Edge Update Policy"
    if (Test-PolicyExists -PolicyName $policyName) {
        Write-LogMessage -Message "Policy '$policyName' already exists, skipping creation" -Type Warning
        return @{ name = $policyName; id = "existing" }
    }
    
    try {
        $body = @{
            name = $policyName
            description = "Microsoft Edge automatic update configuration"
            platforms = "windows10"
            technologies = "mdm"
            settings = @(
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
                }
            )
        }
        
        $result = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies" -Body $body
        Write-LogMessage -Message "Created Edge Update policy" -Type Success
        return $result
    }
    catch {
        Write-LogMessage -Message "Failed to create Edge Update policy - $($_.Exception.Message)" -Type Error
        return $null
    }
}

function New-OfficePolicies {
    Write-LogMessage -Message "Creating Office configuration policy..." -Type Info
    
    $policyName = "Office Updates Configuration"
    if (Test-PolicyExists -PolicyName $policyName) {
        Write-LogMessage -Message "Policy '$policyName' already exists, skipping creation" -Type Warning
        return @{ name = $policyName; id = "existing" }
    }
    
    try {
        $body = @{
            name = $policyName
            description = "Microsoft Office automatic update settings"
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
    
    $policyName = "Outlook Configuration"
    if (Test-PolicyExists -PolicyName $policyName) {
        Write-LogMessage -Message "Policy '$policyName' already exists, skipping creation" -Type Warning
        return @{ name = $policyName; id = "existing" }
    }
    
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
        return $result
    }
    catch {
        Write-LogMessage -Message "Failed to create Outlook policy - $($_.Exception.Message)" -Type Error
        return $null
    }
}

function New-DisableUACPolicy {
    Write-LogMessage -Message "Creating Disable UAC for QuickAssist policy..." -Type Info
    
    $policyName = "Disable UAC for Quickassist"
    if (Test-PolicyExists -PolicyName $policyName) {
        Write-LogMessage -Message "Policy '$policyName' already exists, skipping creation" -Type Warning
        return @{ name = $policyName; id = "existing" }
    }
    
    try {
        $body = @{
            name = $policyName
            description = "Disable UAC secure desktop prompt for QuickAssist remote support"
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
        return $result
    }
    catch {
        Write-LogMessage -Message "Failed to create Disable UAC policy - $($_.Exception.Message)" -Type Error
        return $null
    }
}

# === Helper Functions ===

function Remove-DuplicatePolicies {
    Write-LogMessage -Message "Checking for and removing duplicate policies..." -Type Info
    
    try {
        $existingPolicies = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies" -ErrorAction Stop
        
        # Group policies by name to find duplicates
        $policyGroups = $existingPolicies.value | Group-Object -Property name
        $duplicatesRemoved = 0
        
        foreach ($group in $policyGroups) {
            if ($group.Count -gt 1) {
                Write-LogMessage -Message "Found $($group.Count) policies named '$($group.Name)'" -Type Warning
                
                # Keep the first one, remove the rest
                $policiesToRemove = $group.Group | Select-Object -Skip 1
                
                foreach ($policy in $policiesToRemove) {
                    try {
                        Invoke-MgGraphRequest -Method DELETE -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies/$($policy.id)" -ErrorAction Stop
                        Write-LogMessage -Message "Removed duplicate policy: '$($policy.name)' (ID: $($policy.id))" -Type Success
                        $duplicatesRemoved++
                    }
                    catch {
                        Write-LogMessage -Message "Failed to remove duplicate policy '$($policy.name)': $($_.Exception.Message)" -Type Error
                    }
                }
            }
        }
        
        if ($duplicatesRemoved -gt 0) {
            Write-LogMessage -Message "Removed $duplicatesRemoved duplicate policies" -Type Success
        } else {
            Write-LogMessage -Message "No duplicate policies found" -Type Info
        }
        
        return $duplicatesRemoved
    }
    catch {
        Write-LogMessage -Message "Error checking for duplicate policies: $($_.Exception.Message)" -Type Error
        return 0
    }
}

function Test-PolicyExists {
    param (
        [string]$PolicyName
    )
    
    try {
        Write-LogMessage -Message "Checking if policy '$PolicyName' already exists..." -Type Info -LogOnly
        $existingPolicies = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies" -ErrorAction Stop
        
        # Debug: Log all existing policy names for troubleshooting
        Write-LogMessage -Message "Found $($existingPolicies.value.Count) existing policies in tenant" -Type Info -LogOnly
        foreach ($policy in $existingPolicies.value) {
            Write-LogMessage -Message "Existing policy: '$($policy.name)'" -Type Info -LogOnly
            if ($policy.name -eq $PolicyName) {
                Write-LogMessage -Message "MATCH FOUND: Policy '$PolicyName' already exists with ID: $($policy.id)" -Type Info -LogOnly
                return $true
            }
        }
        Write-LogMessage -Message "NO MATCH: Policy '$PolicyName' does not exist, will create" -Type Info -LogOnly
        return $false
    }
    catch {
        Write-LogMessage -Message "Error checking existing policies: $($_.Exception.Message)" -Type Warning -LogOnly
        return $false
    }
}

function Show-EDREnablementNote {
    Write-LogMessage -Message "=== EDR POLICY SETUP INSTRUCTIONS ===" -Type Warning
    Write-LogMessage -Message "EDR Policy requires manual enablement:" -Type Warning
    Write-LogMessage -Message "1. Go to https://security.microsoft.com" -Type Info
    Write-LogMessage -Message "2. Navigate to Settings > Endpoints > Device management > Onboarding" -Type Info
    Write-LogMessage -Message "3. Enable Microsoft Defender for Business" -Type Info
    Write-LogMessage -Message "4. Configure the security connector" -Type Info
    Write-LogMessage -Message "5. After manual setup, EDR policies will be available" -Type Info
    Write-LogMessage -Message "=====================================" -Type Warning
}

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
            
            Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies/$PolicyId/assignments" -Body $body -ErrorAction Stop
            Write-LogMessage -Message "Assigned policy $PolicyId to $($assignments.Count) groups" -Type Success -LogOnly
            return $true
        }
        else {
            Write-LogMessage -Message "No target groups found for policy $PolicyId assignment" -Type Warning -LogOnly
            return $false
        }
    }
    catch {
        Write-LogMessage -Message "Failed to assign policy $PolicyId - $($_.Exception.Message)" -Type Warning
        return $false
    }
}