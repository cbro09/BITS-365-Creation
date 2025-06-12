#requires -Version 5.1
<#
.SYNOPSIS
    Documentation Module for Microsoft 365 Tenant Setup Utility
.DESCRIPTION
    Generates comprehensive documentation of tenant configuration including groups, policies, SharePoint sites, Intune settings, and users
.NOTES
    Version: 1.0
    Dependencies: Microsoft Graph PowerShell SDK, ImportExcel module
#>

# === Documentation Configuration ===
$DocumentationConfig = @{
    OutputDirectory = "$env:USERPROFILE\Documents\M365TenantSetup_Documentation_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    ReportFormats = @('HTML', 'Excel', 'JSON')
    IncludeScreenshots = $false
    DetailLevel = 'Detailed' # Basic, Standard, Detailed
}

# === Core Documentation Functions ===
function New-TenantDocumentation {
    <#
    .SYNOPSIS
        Main function to generate comprehensive tenant documentation
    .DESCRIPTION
        Creates detailed documentation by populating the Excel template with actual tenant configuration
    #>
    
    try {
        Write-LogMessage -Message "Starting tenant documentation generation..." -Type Info
        
        # Verify Graph connection
        if (-not (Get-MgContext)) {
            Write-LogMessage -Message "Not connected to Microsoft Graph. Please connect first." -Type Error
            return $false
        }
        
        # Create output directory
        $created = New-DocumentationDirectory
        if (-not $created) {
            return $false
        }
        
        # Look for the Excel template
        $templatePath = Find-ExcelTemplate
        if (-not $templatePath) {
            Write-LogMessage -Message "Excel template not found. Please ensure the template file is available." -Type Error
            return $false
        }
        
        # Gather all tenant information
        Write-LogMessage -Message "Gathering tenant configuration data..." -Type Info
        $tenantData = Get-CompleteTenantConfiguration
        
        # Generate populated Excel documentation
        Write-LogMessage -Message "Populating Excel template with configuration data..." -Type Info
        $excelGenerated = New-PopulatedExcelDocumentation -TenantData $tenantData -TemplatePath $templatePath
        
        $documentsGenerated = 0
        if ($excelGenerated) { $documentsGenerated++ }
        
        # Generate supplementary reports
        Write-LogMessage -Message "Generating supplementary HTML report..." -Type Info
        $htmlGenerated = New-HTMLDocumentation -TenantData $tenantData
        if ($htmlGenerated) { $documentsGenerated++ }
        
        # Generate JSON Export for backup
        Write-LogMessage -Message "Generating JSON configuration backup..." -Type Info
        $jsonGenerated = New-JSONDocumentation -TenantData $tenantData
        if ($jsonGenerated) { $documentsGenerated++ }
        
        # Generate Configuration Summary
        Write-LogMessage -Message "Generating configuration summary..." -Type Info
        $summaryGenerated = New-ConfigurationSummary -TenantData $tenantData
        if ($summaryGenerated) { $documentsGenerated++ }
        
        Write-LogMessage -Message "Documentation generation completed. Generated $documentsGenerated documents." -Type Success
        Write-LogMessage -Message "Documentation saved to: $($DocumentationConfig.OutputDirectory)" -Type Info
        
        # Open the documentation directory
        $openDirectory = Read-Host "Would you like to open the documentation directory? (Y/N)"
        if ($openDirectory -eq 'Y' -or $openDirectory -eq 'y') {
            Start-Process -FilePath "explorer.exe" -ArgumentList $DocumentationConfig.OutputDirectory
        }
        
        return $true
    }
    catch {
        Write-LogMessage -Message "Failed to generate documentation: $($_.Exception.Message)" -Type Error
        return $false
    }
}

function New-DocumentationDirectory {
    <#
    .SYNOPSIS
        Creates the documentation output directory structure
    #>
    
    try {
        # Create main directory
        if (-not (Test-Path -Path $DocumentationConfig.OutputDirectory)) {
            New-Item -Path $DocumentationConfig.OutputDirectory -ItemType Directory -Force | Out-Null
        }
        
        # Create subdirectories
        $subDirectories = @('Reports', 'Exports', 'Screenshots', 'Templates')
        foreach ($subDir in $subDirectories) {
            $subDirPath = Join-Path -Path $DocumentationConfig.OutputDirectory -ChildPath $subDir
            if (-not (Test-Path -Path $subDirPath)) {
                New-Item -Path $subDirPath -ItemType Directory -Force | Out-Null
            }
        }
        
        Write-LogMessage -Message "Documentation directory structure created: $($DocumentationConfig.OutputDirectory)" -Type Success -LogOnly
        return $true
    }
    catch {
        Write-LogMessage -Message "Failed to create documentation directory: $($_.Exception.Message)" -Type Error
        return $false
    }
}

function Find-ExcelTemplate {
    <#
    .SYNOPSIS
        Locates the Excel template file for population
    #>
    
    try {
        # Common locations to search for the template
        $searchPaths = @(
            "$env:USERPROFILE\Documents\Master Spreadsheet Customer Details  Test.xlsx",
            "$env:USERPROFILE\Downloads\Master Spreadsheet Customer Details  Test.xlsx",
            ".\Master Spreadsheet Customer Details  Test.xlsx",
            "$env:USERPROFILE\Documents\M365TenantSetup_Documentation*\Templates\Master Spreadsheet Customer Details  Test.xlsx"
        )
        
        foreach ($path in $searchPaths) {
            $resolvedPaths = Resolve-Path -Path $path -ErrorAction SilentlyContinue
            if ($resolvedPaths) {
                foreach ($resolvedPath in $resolvedPaths) {
                    if (Test-Path -Path $resolvedPath) {
                        Write-LogMessage -Message "Found Excel template at: $resolvedPath" -Type Success -LogOnly
                        return $resolvedPath.Path
                    }
                }
            }
        }
        
        # If not found, prompt user to select
        Write-LogMessage -Message "Excel template not found in default locations. Prompting user to select..." -Type Info
        
        Add-Type -AssemblyName System.Windows.Forms
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.Title = "Select Excel Template File"
        $openFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
        $openFileDialog.InitialDirectory = "$env:USERPROFILE\Documents"
        
        if ($openFileDialog.ShowDialog() -eq 'OK') {
            Write-LogMessage -Message "User selected template: $($openFileDialog.FileName)" -Type Success -LogOnly
            return $openFileDialog.FileName
        }
        
        return $null
    }
    catch {
        Write-LogMessage -Message "Error finding Excel template: $($_.Exception.Message)" -Type Error
        return $null
    }
}

function New-PopulatedExcelDocumentation {
    <#
    .SYNOPSIS
        Creates a populated Excel documentation file from the template
    #>
    param (
        [Parameter(Mandatory = $true)]
        [hashtable]$TenantData,
        
        [Parameter(Mandatory = $true)]
        [string]$TemplatePath
    )
    
    try {
        # Create output path
        $outputPath = Join-Path -Path $DocumentationConfig.OutputDirectory -ChildPath "Reports\TenantConfiguration_Populated_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
        
        # Copy template to output location
        Copy-Item -Path $TemplatePath -Destination $outputPath -Force
        
        # Open the Excel file for editing
        $excel = Open-ExcelPackage -Path $outputPath
        
        # Populate each target sheet
        Write-LogMessage -Message "Populating Users sheet..." -Type Info -LogOnly
        Update-UsersSheet -Excel $excel -TenantData $TenantData
        
        Write-LogMessage -Message "Populating Licensing sheet..." -Type Info -LogOnly
        Update-LicensingSheet -Excel $excel -TenantData $TenantData
        
        Write-LogMessage -Message "Populating Hardware Profiles sheet..." -Type Info -LogOnly
        Update-HardwareProfilesSheet -Excel $excel -TenantData $TenantData
        
        Write-LogMessage -Message "Populating Windows Updates sheet..." -Type Info -LogOnly
        Update-WindowsUpdatesSheet -Excel $excel -TenantData $TenantData
        
        Write-LogMessage -Message "Populating Intune Apps sheets..." -Type Info -LogOnly
        Update-IntuneAppsSheets -Excel $excel -TenantData $TenantData
        
        Write-LogMessage -Message "Populating SharePoint Libraries sheet..." -Type Info -LogOnly
        Update-SharePointLibrariesSheet -Excel $excel -TenantData $TenantData
        
        Write-LogMessage -Message "Populating Shared Mailboxes sheet..." -Type Info -LogOnly
        Update-SharedMailboxesSheet -Excel $excel -TenantData $TenantData
        
        Write-LogMessage -Message "Populating Distribution Lists sheet..." -Type Info -LogOnly
        Update-DistributionListsSheet -Excel $excel -TenantData $TenantData
        
        Write-LogMessage -Message "Populating Conditional Access sheet..." -Type Info -LogOnly
        Update-ConditionalAccessSheet -Excel $excel -TenantData $TenantData
        
        # Save and close
        Close-ExcelPackage -ExcelPackage $excel -Save
        
        Write-LogMessage -Message "Populated Excel documentation generated: $outputPath" -Type Success
        return $true
    }
    catch {
        Write-LogMessage -Message "Failed to generate populated Excel documentation: $($_.Exception.Message)" -Type Error
        if ($excel) {
            Close-ExcelPackage -ExcelPackage $excel -NoSave
        }
        return $false
    }
}

# === Sheet Population Functions ===

function Update-UsersSheet {
    <#
    .SYNOPSIS
        Populates the Users sheet with actual user data
    #>
    param (
        [Parameter(Mandatory = $true)]
        $Excel,
        
        [Parameter(Mandatory = $true)]
        [hashtable]$TenantData
    )
    
    try {
        $worksheet = $Excel.Workbook.Worksheets["Users"]
        if (-not $worksheet) {
            Write-LogMessage -Message "Users worksheet not found in template" -Type Warning -LogOnly
            return
        }
        
        $startRow = 7  # Data starts at row 7 based on template structure
        $currentRow = $startRow
        
        foreach ($user in $TenantData.Users.Users) {
            # Extract name parts
            $firstName = Get-SafeString -Value $user.GivenName
            $lastName = Get-SafeString -Value $user.Surname
            
            # Populate user data
            $worksheet.Cells[$currentRow, 1].Value = $firstName                    # Column A: First Name
            $worksheet.Cells[$currentRow, 2].Value = $lastName                     # Column B: Last Name
            $worksheet.Cells[$currentRow, 3].Value = $user.UserPrincipalName       # Column C: Email
            $worksheet.Cells[$currentRow, 4].Value = Get-SafeString -Value $user.JobTitle        # Column D: Job Title
            # Manager email would need to be resolved from manager ID
            $worksheet.Cells[$currentRow, 6].Value = Get-SafeString -Value $user.Department      # Column F: Department
            $worksheet.Cells[$currentRow, 7].Value = Get-SafeString -Value $user.Office          # Column G: Office location
            # Phone number would be in additional properties
            
            $currentRow++
            
            # Limit to prevent performance issues
            if ($currentRow > $startRow + 500) {
                Write-LogMessage -Message "Limited users export to first 500 users" -Type Warning -LogOnly
                break
            }
        }
        
        Write-LogMessage -Message "Updated Users sheet with $($currentRow - $startRow) users" -Type Success -LogOnly
    }
    catch {
        Write-LogMessage -Message "Error updating Users sheet: $($_.Exception.Message)" -Type Warning -LogOnly
    }
}

function Update-LicensingSheet {
    <#
    .SYNOPSIS
        Populates the Licensing sheet with user license assignments
    #>
    param (
        [Parameter(Mandatory = $true)]
        $Excel,
        
        [Parameter(Mandatory = $true)]
        [hashtable]$TenantData
    )
    
    try {
        $worksheet = $Excel.Workbook.Worksheets["Licensing"]
        if (-not $worksheet) {
            Write-LogMessage -Message "Licensing worksheet not found in template" -Type Warning -LogOnly
            return
        }
        
        $startRow = 8  # Data starts at row 8 based on template structure
        $currentRow = $startRow
        
        foreach ($user in $TenantData.Users.Users) {
            if ($user.AssignedLicenses -and $user.AssignedLicenses.Count -gt 0) {
                $worksheet.Cells[$currentRow, 2].Value = $user.DisplayName         # Column B: User Name
                $worksheet.Cells[$currentRow, 3].Value = $user.AssignedLicenses[0] # Column C: Base License Type
                
                # Additional licenses in subsequent columns
                if ($user.AssignedLicenses.Count -gt 1) {
                    $worksheet.Cells[$currentRow, 4].Value = $user.AssignedLicenses[1]  # Column D: Additional Software 1
                }
                if ($user.AssignedLicenses.Count -gt 2) {
                    $worksheet.Cells[$currentRow, 5].Value = $user.AssignedLicenses[2]  # Column E: Additional Software 2
                }
                
                $currentRow++
            }
            
            # Limit to prevent performance issues
            if ($currentRow > $startRow + 500) {
                break
            }
        }
        
        Write-LogMessage -Message "Updated Licensing sheet with license assignments" -Type Success -LogOnly
    }
    catch {
        Write-LogMessage -Message "Error updating Licensing sheet: $($_.Exception.Message)" -Type Warning -LogOnly
    }
}

function Update-HardwareProfilesSheet {
    <#
    .SYNOPSIS
        Populates the Hardware Profiles sheet with Intune policies
    #>
    param (
        [Parameter(Mandatory = $true)]
        $Excel,
        
        [Parameter(Mandatory = $true)]
        [hashtable]$TenantData
    )
    
    try {
        $worksheet = $Excel.Workbook.Worksheets["Hardware Profiles"]
        if (-not $worksheet) {
            Write-LogMessage -Message "Hardware Profiles worksheet not found in template" -Type Warning -LogOnly
            return
        }
        
        # Configuration Policies section (starts around row 10)
        $configRow = 11  # Row for configuration policy names
        $configCol = 2   # Column B
        
        if ($TenantData.Intune.DeviceConfigurationPolicies.Count -gt 0) {
            foreach ($policy in $TenantData.Intune.DeviceConfigurationPolicies) {
                $worksheet.Cells[$configRow, $configCol].Value = $policy.DisplayName
                $configRow++
                
                # Limit entries
                if ($configRow > 20) { break }
            }
        } else {
            $worksheet.Cells[11, $configCol].Value = "No configuration policies found"
        }
        
        # Compliance Policies section (starts around row 21)
        $complianceRow = 22  # Row for compliance policy names
        $complianceCol = 2   # Column B
        
        if ($TenantData.Intune.DeviceCompliancePolicies.Count -gt 0) {
            foreach ($policy in $TenantData.Intune.DeviceCompliancePolicies) {
                $worksheet.Cells[$complianceRow, $complianceCol].Value = $policy.DisplayName
                $complianceRow++
                
                # Limit entries
                if ($complianceRow > 30) { break }
            }
        } else {
            $worksheet.Cells[22, $complianceCol].Value = "No compliance policies found"
        }
        
        Write-LogMessage -Message "Updated Hardware Profiles sheet with Intune policies" -Type Success -LogOnly
    }
    catch {
        Write-LogMessage -Message "Error updating Hardware Profiles sheet: $($_.Exception.Message)" -Type Warning -LogOnly
    }
}

function Update-ConditionalAccessSheet {
    <#
    .SYNOPSIS
        Populates the Conditional Access sheet with actual CA policies
    #>
    param (
        [Parameter(Mandatory = $true)]
        $Excel,
        
        [Parameter(Mandatory = $true)]
        [hashtable]$TenantData
    )
    
    try {
        $worksheet = $Excel.Workbook.Worksheets["Conditional Access"]
        if (-not $worksheet) {
            Write-LogMessage -Message "Conditional Access worksheet not found in template" -Type Warning -LogOnly
            return
        }
        
        $startRow = 13  # Start after the template entries
        $currentRow = $startRow
        
        # Add actual policies that were created
        foreach ($policy in $TenantData.ConditionalAccess.Policies) {
            $worksheet.Cells[$currentRow, 2].Value = $policy.DisplayName  # Column B: Policy Name
            
            # Create policy setting description
            $policyDescription = "State: $($policy.State)"
            if ($policy.Conditions.Users.IncludeUsers) {
                $policyDescription += " | Users: $($policy.Conditions.Users.IncludeUsers -join ', ')"
            }
            if ($policy.Conditions.Applications.IncludeApplications) {
                $policyDescription += " | Apps: $($policy.Conditions.Applications.IncludeApplications -join ', ')"
            }
            if ($policy.GrantControls.BuiltInControls) {
                $policyDescription += " | Grant: $($policy.GrantControls.BuiltInControls -join ', ')"
            }
            
            $worksheet.Cells[$currentRow, 4].Value = $policyDescription  # Column D: Policy Setting
            $currentRow++
            
            # Limit entries
            if ($currentRow > $startRow + 20) { break }
        }
        
        Write-LogMessage -Message "Updated Conditional Access sheet with $($currentRow - $startRow) policies" -Type Success -LogOnly
    }
    catch {
        Write-LogMessage -Message "Error updating Conditional Access sheet: $($_.Exception.Message)" -Type Warning -LogOnly
    }
}

function Update-SharePointLibrariesSheet {
    <#
    .SYNOPSIS
        Populates the SharePoint Libraries sheet
    #>
    param (
        [Parameter(Mandatory = $true)]
        $Excel,
        
        [Parameter(Mandatory = $true)]
        [hashtable]$TenantData
    )
    
    try {
        $worksheet = $Excel.Workbook.Worksheets["SharePoint Libaries"]  # Note: matches template spelling
        if (-not $worksheet) {
            Write-LogMessage -Message "SharePoint Libraries worksheet not found in template" -Type Warning -LogOnly
            return
        }
        
        $startRow = 13  # Start after template entries
        $currentRow = $startRow
        
        foreach ($site in $TenantData.SharePoint.SiteCollections) {
            $worksheet.Cells[$currentRow, 2].Value = $site.DisplayName  # Column B: Site Name
            $worksheet.Cells[$currentRow, 4].Value = "Site Admin"       # Column D: Approver
            $currentRow++
            
            # Limit entries
            if ($currentRow > $startRow + 20) { break }
        }
        
        Write-LogMessage -Message "Updated SharePoint Libraries sheet" -Type Success -LogOnly
    }
    catch {
        Write-LogMessage -Message "Error updating SharePoint Libraries sheet: $($_.Exception.Message)" -Type Warning -LogOnly
    }
}

function Update-WindowsUpdatesSheet {
    <#
    .SYNOPSIS
        Populates the Windows Updates sheet
    #>
    param (
        [Parameter(Mandatory = $true)]
        $Excel,
        
        [Parameter(Mandatory = $true)]
        [hashtable]$TenantData
    )
    
    try {
        $worksheet = $Excel.Workbook.Worksheets["Windows Updates"]
        if (-not $worksheet) {
            Write-LogMessage -Message "Windows Updates worksheet not found in template" -Type Warning -LogOnly
            return
        }
        
        # Add placeholder or actual update ring data if available
        $worksheet.Cells[10, 2].Value = "Standard Update Ring Configured"
        Write-LogMessage -Message "Updated Windows Updates sheet" -Type Success -LogOnly
    }
    catch {
        Write-LogMessage -Message "Error updating Windows Updates sheet: $($_.Exception.Message)" -Type Warning -LogOnly
    }
}

function Update-IntuneAppsSheets {
    <#
    .SYNOPSIS
        Populates all Intune Apps sheets
    #>
    param (
        [Parameter(Mandatory = $true)]
        $Excel,
        
        [Parameter(Mandatory = $true)]
        [hashtable]$TenantData
    )
    
    try {
        $appSheets = @(
            "Intune Windows Apps",
            "Intune Android Apps", 
            "Intune Apple IOS Apps",
            "Intune Apple iPadOS Apps",
            "Intune Mac OS Apps"
        )
        
        foreach ($sheetName in $appSheets) {
            $worksheet = $Excel.Workbook.Worksheets[$sheetName]
            if ($worksheet) {
                # Add placeholder indicating apps are configured
                $worksheet.Cells[15, 2].Value = "Apps configured via Intune"
                Write-LogMessage -Message "Updated $sheetName sheet" -Type Success -LogOnly
            }
        }
    }
    catch {
        Write-LogMessage -Message "Error updating Intune Apps sheets: $($_.Exception.Message)" -Type Warning -LogOnly
    }
}

function Update-SharedMailboxesSheet {
    <#
    .SYNOPSIS
        Populates the Shared Mailboxes sheet
    #>
    param (
        [Parameter(Mandatory = $true)]
        $Excel,
        
        [Parameter(Mandatory = $true)]
        [hashtable]$TenantData
    )
    
    try {
        $worksheet = $Excel.Workbook.Worksheets["Shared Mailboxes"]
        if (-not $worksheet) {
            Write-LogMessage -Message "Shared Mailboxes worksheet not found in template" -Type Warning -LogOnly
            return
        }
        
        # Add shared mailboxes from groups if any are mail-enabled
        $startRow = 8
        $currentRow = $startRow
        
        $sharedMailboxes = $TenantData.Groups.DistributionGroups | Where-Object { $_.DisplayName -like "*shared*" -or $_.DisplayName -like "*mailbox*" }
        
        foreach ($mailbox in $sharedMailboxes) {
            $worksheet.Cells[$currentRow, 2].Value = $mailbox.DisplayName
            $currentRow++
            
            if ($currentRow > $startRow + 10) { break }
        }
        
        Write-LogMessage -Message "Updated Shared Mailboxes sheet" -Type Success -LogOnly
    }
    catch {
        Write-LogMessage -Message "Error updating Shared Mailboxes sheet: $($_.Exception.Message)" -Type Warning -LogOnly
    }
}

function Update-DistributionListsSheet {
    <#
    .SYNOPSIS
        Populates the Distribution Lists sheet
    #>
    param (
        [Parameter(Mandatory = $true)]
        $Excel,
        
        [Parameter(Mandatory = $true)]
        [hashtable]$TenantData
    )
    
    try {
        $worksheet = $Excel.Workbook.Worksheets["Distribution list"]
        if (-not $worksheet) {
            Write-LogMessage -Message "Distribution Lists worksheet not found in template" -Type Warning -LogOnly
            return
        }
        
        $startRow = 8
        $currentRow = $startRow
        
        foreach ($group in $TenantData.Groups.DistributionGroups) {
            $worksheet.Cells[$currentRow, 2].Value = $group.DisplayName        # Group Name
            $worksheet.Cells[$currentRow, 3].Value = $group.Description        # Description
            $worksheet.Cells[$currentRow, 4].Value = $group.MemberCount        # Member Count
            $currentRow++
            
            if ($currentRow > $startRow + 20) { break }
        }
        
        Write-LogMessage -Message "Updated Distribution Lists sheet with $($currentRow - $startRow) groups" -Type Success -LogOnly
    }
    catch {
        Write-LogMessage -Message "Error updating Distribution Lists sheet: $($_.Exception.Message)" -Type Warning -LogOnly
    }
}
    <#
    .SYNOPSIS
        Gathers comprehensive tenant configuration data
    #>
    
    Write-LogMessage -Message "Gathering tenant configuration data..." -Type Info
    
    $tenantData = @{
        GeneratedOn = Get-Date
        TenantInfo = @{}
        Groups = @{}
        ConditionalAccess = @{}
        SharePoint = @{}
        Intune = @{}
        Users = @{}
        Licenses = @{}
        Security = @{}
        Compliance = @{}
    }
    
    try {
        # Basic Tenant Information
        Write-LogMessage -Message "Collecting tenant information..." -Type Info -LogOnly
        $tenantData.TenantInfo = Get-TenantInformation
        
        # Groups Information
        Write-LogMessage -Message "Collecting groups information..." -Type Info -LogOnly
        $tenantData.Groups = Get-GroupsInformation
        
        # Conditional Access Policies
        Write-LogMessage -Message "Collecting conditional access policies..." -Type Info -LogOnly
        $tenantData.ConditionalAccess = Get-ConditionalAccessInformation
        
        # SharePoint Information
        Write-LogMessage -Message "Collecting SharePoint information..." -Type Info -LogOnly
        $tenantData.SharePoint = Get-SharePointInformation
        
        # Intune Information
        Write-LogMessage -Message "Collecting Intune information..." -Type Info -LogOnly
        $tenantData.Intune = Get-IntuneInformation
        
        # Users Information
        Write-LogMessage -Message "Collecting users information..." -Type Info -LogOnly
        $tenantData.Users = Get-UsersInformation
        
        # License Information
        Write-LogMessage -Message "Collecting license information..." -Type Info -LogOnly
        $tenantData.Licenses = Get-LicenseInformation
        
        # Security Settings
        Write-LogMessage -Message "Collecting security settings..." -Type Info -LogOnly
        $tenantData.Security = Get-SecurityInformation
        
        Write-LogMessage -Message "Tenant configuration data collection completed" -Type Success -LogOnly
        return $tenantData
    }
    catch {
        Write-LogMessage -Message "Error collecting tenant data: $($_.Exception.Message)" -Type Error
        return $tenantData
    }
}

function Get-TenantInformation {
    <#
    .SYNOPSIS
        Collects basic tenant information
    #>
    
    try {
        $organization = Get-MgOrganization
        $context = Get-MgContext
        
        $tenantInfo = @{
            TenantId = $organization.Id
            DisplayName = $organization.DisplayName
            DefaultDomain = ($organization.VerifiedDomains | Where-Object { $_.IsDefault -eq $true }).Name
            VerifiedDomains = $organization.VerifiedDomains | ForEach-Object { $_.Name }
            CountryCode = $organization.CountryLetterCode
            City = $organization.City
            State = $organization.State
            CreatedDateTime = $organization.CreatedDateTime
            ConnectedAs = $context.Account
            ConnectedScopes = $context.Scopes
            SetupDate = Get-Date
        }
        
        # Add script tenant state if available
        if ($script:TenantState) {
            $tenantInfo.AdminEmail = $script:TenantState.AdminEmail
            $tenantInfo.SetupGroups = $script:TenantState.CreatedGroups
        }
        
        return $tenantInfo
    }
    catch {
        Write-LogMessage -Message "Error collecting tenant information: $($_.Exception.Message)" -Type Warning -LogOnly
        return @{}
    }
}

function Get-GroupsInformation {
    <#
    .SYNOPSIS
        Collects information about all groups in the tenant
    #>
    
    try {
        $groups = Get-MgGroup -All
        $groupsInfo = @{
            SecurityGroups = @()
            DistributionGroups = @()
            Microsoft365Groups = @()
            DynamicGroups = @()
            TotalCount = $groups.Count
        }
        
        foreach ($group in $groups) {
            $groupData = @{
                Id = $group.Id
                DisplayName = $group.DisplayName
                Description = $group.Description
                GroupTypes = $group.GroupTypes
                SecurityEnabled = $group.SecurityEnabled
                MailEnabled = $group.MailEnabled
                CreatedDateTime = $group.CreatedDateTime
                MembershipRule = $group.MembershipRule
                MembershipRuleProcessingState = $group.MembershipRuleProcessingState
            }
            
            # Try to get member count
            try {
                $members = Get-MgGroupMember -GroupId $group.Id -All
                $groupData.MemberCount = $members.Count
            }
            catch {
                $groupData.MemberCount = "Unable to retrieve"
            }
            
            # Categorize groups
            if ($group.GroupTypes -contains "Unified") {
                $groupsInfo.Microsoft365Groups += $groupData
            }
            elseif ($group.GroupTypes -contains "DynamicMembership") {
                $groupsInfo.DynamicGroups += $groupData
            }
            elseif ($group.SecurityEnabled -and -not $group.MailEnabled) {
                $groupsInfo.SecurityGroups += $groupData
            }
            elseif ($group.MailEnabled -and -not $group.SecurityEnabled) {
                $groupsInfo.DistributionGroups += $groupData
            }
        }
        
        return $groupsInfo
    }
    catch {
        Write-LogMessage -Message "Error collecting groups information: $($_.Exception.Message)" -Type Warning -LogOnly
        return @{}
    }
}

function Get-ConditionalAccessInformation {
    <#
    .SYNOPSIS
        Collects conditional access policies information
    #>
    
    try {
        $policies = Get-MgIdentityConditionalAccessPolicy -All
        $caInfo = @{
            Policies = @()
            TotalCount = $policies.Count
            EnabledCount = ($policies | Where-Object { $_.State -eq "enabled" }).Count
            DisabledCount = ($policies | Where-Object { $_.State -eq "disabled" }).Count
        }
        
        foreach ($policy in $policies) {
            $policyData = @{
                Id = $policy.Id
                DisplayName = $policy.DisplayName
                State = $policy.State
                CreatedDateTime = $policy.CreatedDateTime
                ModifiedDateTime = $policy.ModifiedDateTime
                Conditions = @{
                    Users = $policy.Conditions.Users
                    Applications = $policy.Conditions.Applications
                    Platforms = $policy.Conditions.Platforms
                    Locations = $policy.Conditions.Locations
                    ClientAppTypes = $policy.Conditions.ClientAppTypes
                    SignInRiskLevels = $policy.Conditions.SignInRiskLevels
                    UserRiskLevels = $policy.Conditions.UserRiskLevels
                }
                GrantControls = $policy.GrantControls
                SessionControls = $policy.SessionControls
            }
            
            $caInfo.Policies += $policyData
        }
        
        return $caInfo
    }
    catch {
        Write-LogMessage -Message "Error collecting conditional access information: $($_.Exception.Message)" -Type Warning -LogOnly
        return @{}
    }
}

function Get-SharePointInformation {
    <#
    .SYNOPSIS
        Collects SharePoint configuration information
    #>
    
    try {
        # This would require SharePoint Online PowerShell module
        # For now, return placeholder structure
        $spInfo = @{
            TenantSettings = @{}
            SiteCollections = @()
            TotalSites = 0
            StorageUsed = "Not available"
            SharingSettings = "Not available"
            ExternalSharingEnabled = "Not available"
        }
        
        # Try to get basic SharePoint info through Graph if available
        try {
            $sites = Get-MgSite -All -Top 50
            $spInfo.TotalSites = $sites.Count
            $spInfo.SiteCollections = $sites | ForEach-Object {
                @{
                    Id = $_.Id
                    DisplayName = $_.DisplayName
                    WebUrl = $_.WebUrl
                    CreatedDateTime = $_.CreatedDateTime
                    LastModifiedDateTime = $_.LastModifiedDateTime
                }
            }
        }
        catch {
            Write-LogMessage -Message "SharePoint sites information not available through Graph API" -Type Warning -LogOnly
        }
        
        return $spInfo
    }
    catch {
        Write-LogMessage -Message "Error collecting SharePoint information: $($_.Exception.Message)" -Type Warning -LogOnly
        return @{}
    }
}

function Get-IntuneInformation {
    <#
    .SYNOPSIS
        Collects Intune configuration information
    #>
    
    try {
        $intuneInfo = @{
            DeviceCompliancePolicies = @()
            DeviceConfigurationPolicies = @()
            AppProtectionPolicies = @()
            EnrollmentRestrictions = @()
            ManagedDevices = @()
            TotalDevices = 0
        }
        
        # Device Compliance Policies
        try {
            $compliancePolicies = Get-MgDeviceManagementDeviceCompliancePolicy -All
            $intuneInfo.DeviceCompliancePolicies = $compliancePolicies | ForEach-Object {
                @{
                    Id = $_.Id
                    DisplayName = $_.DisplayName
                    Description = $_.Description
                    CreatedDateTime = $_.CreatedDateTime
                    LastModifiedDateTime = $_.LastModifiedDateTime
                    Version = $_.Version
                }
            }
        }
        catch {
            Write-LogMessage -Message "Could not retrieve device compliance policies" -Type Warning -LogOnly
        }
        
        # Device Configuration Policies
        try {
            $configPolicies = Get-MgDeviceManagementDeviceConfiguration -All
            $intuneInfo.DeviceConfigurationPolicies = $configPolicies | ForEach-Object {
                @{
                    Id = $_.Id
                    DisplayName = $_.DisplayName
                    Description = $_.Description
                    CreatedDateTime = $_.CreatedDateTime
                    LastModifiedDateTime = $_.LastModifiedDateTime
                    Version = $_.Version
                }
            }
        }
        catch {
            Write-LogMessage -Message "Could not retrieve device configuration policies" -Type Warning -LogOnly
        }
        
        # Managed Devices
        try {
            $devices = Get-MgDeviceManagementManagedDevice -All -Top 100
            $intuneInfo.TotalDevices = $devices.Count
            $intuneInfo.ManagedDevices = $devices | ForEach-Object {
                @{
                    Id = $_.Id
                    DeviceName = $_.DeviceName
                    OperatingSystem = $_.OperatingSystem
                    OSVersion = $_.OSVersion
                    ComplianceState = $_.ComplianceState
                    EnrolledDateTime = $_.EnrolledDateTime
                    LastSyncDateTime = $_.LastSyncDateTime
                    UserPrincipalName = $_.UserPrincipalName
                }
            }
        }
        catch {
            Write-LogMessage -Message "Could not retrieve managed devices" -Type Warning -LogOnly
        }
        
        return $intuneInfo
    }
    catch {
        Write-LogMessage -Message "Error collecting Intune information: $($_.Exception.Message)" -Type Warning -LogOnly
        return @{}
    }
}

function Get-UsersInformation {
    <#
    .SYNOPSIS
        Collects users information
    #>
    
    try {
        $users = Get-MgUser -All -Top 500
        $usersInfo = @{
            TotalUsers = $users.Count
            EnabledUsers = ($users | Where-Object { $_.AccountEnabled -eq $true }).Count
            DisabledUsers = ($users | Where-Object { $_.AccountEnabled -eq $false }).Count
            GuestUsers = ($users | Where-Object { $_.UserType -eq "Guest" }).Count
            Users = @()
        }
        
        foreach ($user in $users) {
            $userData = @{
                Id = $user.Id
                UserPrincipalName = $user.UserPrincipalName
                DisplayName = $user.DisplayName
                GivenName = $user.GivenName
                Surname = $user.Surname
                JobTitle = $user.JobTitle
                Department = $user.Department
                Office = $user.OfficeLocation
                AccountEnabled = $user.AccountEnabled
                UserType = $user.UserType
                CreatedDateTime = $user.CreatedDateTime
                SignInActivity = "Not available"
            }
            
            # Try to get license information
            try {
                $userLicenses = Get-MgUserLicenseDetail -UserId $user.Id
                $userData.AssignedLicenses = $userLicenses | ForEach-Object { $_.SkuPartNumber }
            }
            catch {
                $userData.AssignedLicenses = @()
            }
            
            $usersInfo.Users += $userData
        }
        
        return $usersInfo
    }
    catch {
        Write-LogMessage -Message "Error collecting users information: $($_.Exception.Message)" -Type Warning -LogOnly
        return @{}
    }
}

function Get-LicenseInformation {
    <#
    .SYNOPSIS
        Collects license information
    #>
    
    try {
        $subscribedSkus = Get-MgSubscribedSku
        $licenseInfo = @{
            SubscribedSkus = @()
            TotalLicenses = 0
            UsedLicenses = 0
        }
        
        foreach ($sku in $subscribedSkus) {
            $skuData = @{
                SkuId = $sku.SkuId
                SkuPartNumber = $sku.SkuPartNumber
                ServicePlans = $sku.ServicePlans | ForEach-Object { @{ ServicePlanName = $_.ServicePlanName; ServicePlanId = $_.ServicePlanId } }
                PrepaidUnits = $sku.PrepaidUnits
                ConsumedUnits = $sku.ConsumedUnits
                CapabilityStatus = $sku.CapabilityStatus
            }
            
            $licenseInfo.SubscribedSkus += $skuData
            $licenseInfo.TotalLicenses += $sku.PrepaidUnits.Enabled
            $licenseInfo.UsedLicenses += $sku.ConsumedUnits
        }
        
        return $licenseInfo
    }
    catch {
        Write-LogMessage -Message "Error collecting license information: $($_.Exception.Message)" -Type Warning -LogOnly
        return @{}
    }
}

function Get-SecurityInformation {
    <#
    .SYNOPSIS
        Collects security settings information
    #>
    
    try {
        $securityInfo = @{
            SecurityDefaults = @{}
            PasswordPolicy = @{}
            MFAStatus = "Not available"
            RiskyUsers = 0
            RiskySignIns = 0
        }
        
        # Try to get security defaults status
        try {
            $securityDefaults = Get-MgPolicyIdentitySecurityDefaultEnforcementPolicy
            $securityInfo.SecurityDefaults = @{
                IsEnabled = $securityDefaults.IsEnabled
                Description = $securityDefaults.Description
            }
        }
        catch {
            Write-LogMessage -Message "Could not retrieve security defaults status" -Type Warning -LogOnly
        }
        
        return $securityInfo
    }
    catch {
        Write-LogMessage -Message "Error collecting security information: $($_.Exception.Message)" -Type Warning -LogOnly
        return @{}
    }
}

function New-HTMLDocumentation {
    <#
    .SYNOPSIS
        Generates HTML documentation report
    #>
    param (
        [Parameter(Mandatory = $true)]
        [hashtable]$TenantData
    )
    
    try {
        $htmlPath = Join-Path -Path $DocumentationConfig.OutputDirectory -ChildPath "Reports\TenantConfiguration.html"
        
        $htmlContent = @"
<!DOCTYPE html>
<html>
<head>
    <title>Microsoft 365 Tenant Configuration Report</title>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 20px; background-color: #f5f5f5; }
        .container { max-width: 1200px; margin: 0 auto; background-color: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        h1 { color: #0078d4; border-bottom: 3px solid #0078d4; padding-bottom: 10px; }
        h2 { color: #106ebe; border-left: 4px solid #0078d4; padding-left: 10px; margin-top: 30px; }
        h3 { color: #323130; }
        .info-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 20px; margin: 20px 0; }
        .info-card { background-color: #f8f9fa; border: 1px solid #e9ecef; border-radius: 6px; padding: 15px; }
        .info-card h4 { margin-top: 0; color: #0078d4; }
        table { width: 100%; border-collapse: collapse; margin: 15px 0; }
        th, td { border: 1px solid #ddd; padding: 12px; text-align: left; }
        th { background-color: #0078d4; color: white; }
        tr:nth-child(even) { background-color: #f2f2f2; }
        .status-enabled { color: #107c10; font-weight: bold; }
        .status-disabled { color: #d13438; font-weight: bold; }
        .timestamp { color: #666; font-size: 0.9em; }
        .summary-stats { display: flex; justify-content: space-around; margin: 20px 0; }
        .stat-box { text-align: center; padding: 15px; background-color: #e3f2fd; border-radius: 6px; }
        .stat-number { font-size: 2em; font-weight: bold; color: #0078d4; }
        .stat-label { color: #666; }
    </style>
</head>
<body>
    <div class="container">
        <h1>Microsoft 365 Tenant Configuration Report</h1>
        <p class="timestamp">Generated on: $($TenantData.GeneratedOn.ToString('yyyy-MM-dd HH:mm:ss'))</p>
        
        <h2>Tenant Information</h2>
        <div class="info-grid">
            <div class="info-card">
                <h4>Basic Information</h4>
                <p><strong>Tenant Name:</strong> $($TenantData.TenantInfo.DisplayName)</p>
                <p><strong>Default Domain:</strong> $($TenantData.TenantInfo.DefaultDomain)</p>
                <p><strong>Tenant ID:</strong> $($TenantData.TenantInfo.TenantId)</p>
                <p><strong>Country:</strong> $($TenantData.TenantInfo.CountryCode)</p>
            </div>
            <div class="info-card">
                <h4>Setup Information</h4>
                <p><strong>Connected As:</strong> $($TenantData.TenantInfo.ConnectedAs)</p>
                <p><strong>Setup Date:</strong> $($TenantData.TenantInfo.SetupDate)</p>
                <p><strong>Admin Email:</strong> $($TenantData.TenantInfo.AdminEmail)</p>
            </div>
        </div>
        
        <h2>Summary Statistics</h2>
        <div class="summary-stats">
            <div class="stat-box">
                <div class="stat-number">$($TenantData.Groups.TotalCount)</div>
                <div class="stat-label">Total Groups</div>
            </div>
            <div class="stat-box">
                <div class="stat-number">$($TenantData.Users.TotalUsers)</div>
                <div class="stat-label">Total Users</div>
            </div>
            <div class="stat-box">
                <div class="stat-number">$($TenantData.ConditionalAccess.TotalCount)</div>
                <div class="stat-label">CA Policies</div>
            </div>
            <div class="stat-box">
                <div class="stat-number">$($TenantData.Intune.TotalDevices)</div>
                <div class="stat-label">Managed Devices</div>
            </div>
        </div>
        
        <h2>Groups Configuration</h2>
        <div class="info-grid">
            <div class="info-card">
                <h4>Group Distribution</h4>
                <p><strong>Security Groups:</strong> $($TenantData.Groups.SecurityGroups.Count)</p>
                <p><strong>Microsoft 365 Groups:</strong> $($TenantData.Groups.Microsoft365Groups.Count)</p>
                <p><strong>Distribution Groups:</strong> $($TenantData.Groups.DistributionGroups.Count)</p>
                <p><strong>Dynamic Groups:</strong> $($TenantData.Groups.DynamicGroups.Count)</p>
            </div>
        </div>
        
        <h3>Security Groups</h3>
        <table>
            <tr><th>Name</th><th>Description</th><th>Member Count</th><th>Created</th></tr>
"@

        foreach ($group in $TenantData.Groups.SecurityGroups) {
            $htmlContent += @"
            <tr>
                <td>$($group.DisplayName)</td>
                <td>$(Get-SafeString -Value $group.Description -MaxLength 100)</td>
                <td>$($group.MemberCount)</td>
                <td>$(if ($group.CreatedDateTime) { [DateTime]$group.CreatedDateTime | Get-Date -Format 'yyyy-MM-dd' } else { 'N/A' })</td>
            </tr>
"@
        }

        $htmlContent += @"
        </table>
        
        <h2>Conditional Access Policies</h2>
        <p><strong>Total Policies:</strong> $($TenantData.ConditionalAccess.TotalCount) | 
           <span class="status-enabled">Enabled: $($TenantData.ConditionalAccess.EnabledCount)</span> | 
           <span class="status-disabled">Disabled: $($TenantData.ConditionalAccess.DisabledCount)</span></p>
        
        <table>
            <tr><th>Name</th><th>State</th><th>Created</th><th>Modified</th></tr>
"@

        foreach ($policy in $TenantData.ConditionalAccess.Policies) {
            $stateClass = if ($policy.State -eq "enabled") { "status-enabled" } else { "status-disabled" }
            $htmlContent += @"
            <tr>
                <td>$($policy.DisplayName)</td>
                <td><span class="$stateClass">$($policy.State)</span></td>
                <td>$(if ($policy.CreatedDateTime) { [DateTime]$policy.CreatedDateTime | Get-Date -Format 'yyyy-MM-dd' } else { 'N/A' })</td>
                <td>$(if ($policy.ModifiedDateTime) { [DateTime]$policy.ModifiedDateTime | Get-Date -Format 'yyyy-MM-dd' } else { 'N/A' })</td>
            </tr>
"@
        }

        $htmlContent += @"
        </table>
        
        <h2>License Information</h2>
        <table>
            <tr><th>License</th><th>Total</th><th>Used</th><th>Available</th><th>Usage %</th></tr>
"@

        foreach ($sku in $TenantData.Licenses.SubscribedSkus) {
            $total = $sku.PrepaidUnits.Enabled
            $used = $sku.ConsumedUnits
            $available = $total - $used
            $usagePercent = if ($total -gt 0) { [math]::Round(($used / $total) * 100, 1) } else { 0 }
            
            $htmlContent += @"
            <tr>
                <td>$($sku.SkuPartNumber)</td>
                <td>$total</td>
                <td>$used</td>
                <td>$available</td>
                <td>$usagePercent%</td>
            </tr>
"@
        }

        $htmlContent += @"
        </table>
        
        <h2>Users Summary</h2>
        <div class="info-grid">
            <div class="info-card">
                <h4>User Statistics</h4>
                <p><strong>Total Users:</strong> $($TenantData.Users.TotalUsers)</p>
                <p><strong>Enabled Users:</strong> $($TenantData.Users.EnabledUsers)</p>
                <p><strong>Disabled Users:</strong> $($TenantData.Users.DisabledUsers)</p>
                <p><strong>Guest Users:</strong> $($TenantData.Users.GuestUsers)</p>
            </div>
        </div>
        
        <footer style="margin-top: 40px; text-align: center; color: #666; border-top: 1px solid #ddd; padding-top: 20px;">
            <p>Generated by Microsoft 365 Tenant Setup Utility | $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
        </footer>
    </div>
</body>
</html>
"@

        $htmlContent | Out-File -FilePath $htmlPath -Encoding UTF8
        Write-LogMessage -Message "HTML documentation generated: $htmlPath" -Type Success -LogOnly
        return $true
    }
    catch {
        Write-LogMessage -Message "Failed to generate HTML documentation: $($_.Exception.Message)" -Type Error
        return $false
    }
}

function New-ExcelDocumentation {
    <#
    .SYNOPSIS
        Generates basic Excel documentation report as backup
    #>
    param (
        [Parameter(Mandatory = $true)]
        [hashtable]$TenantData
    )
    
    try {
        $excelPath = Join-Path -Path $DocumentationConfig.OutputDirectory -ChildPath "Reports\TenantConfiguration_BasicReport.xlsx"
        
        # Summary sheet
        $summaryData = @(
            [PSCustomObject]@{ Property = "Tenant Name"; Value = $TenantData.TenantInfo.DisplayName }
            [PSCustomObject]@{ Property = "Default Domain"; Value = $TenantData.TenantInfo.DefaultDomain }
            [PSCustomObject]@{ Property = "Tenant ID"; Value = $TenantData.TenantInfo.TenantId }
            [PSCustomObject]@{ Property = "Setup Date"; Value = $TenantData.TenantInfo.SetupDate }
            [PSCustomObject]@{ Property = "Connected As"; Value = $TenantData.TenantInfo.ConnectedAs }
            [PSCustomObject]@{ Property = "Total Groups"; Value = $TenantData.Groups.TotalCount }
            [PSCustomObject]@{ Property = "Total Users"; Value = $TenantData.Users.TotalUsers }
            [PSCustomObject]@{ Property = "CA Policies"; Value = $TenantData.ConditionalAccess.TotalCount }
            [PSCustomObject]@{ Property = "Managed Devices"; Value = $TenantData.Intune.TotalDevices }
        )
        
        # Export to Excel with multiple sheets
        $summaryData | Export-Excel -Path $excelPath -WorksheetName "Summary" -AutoSize -BoldTopRow
        
        # Groups sheet
        if ($TenantData.Groups.SecurityGroups.Count -gt 0) {
            $TenantData.Groups.SecurityGroups | Export-Excel -Path $excelPath -WorksheetName "Security Groups" -AutoSize -BoldTopRow
        }
        
        # Users sheet (first 1000 users to avoid Excel limits)
        if ($TenantData.Users.Users.Count -gt 0) {
            $TenantData.Users.Users | Select-Object -First 1000 | Export-Excel -Path $excelPath -WorksheetName "Users" -AutoSize -BoldTopRow
        }
        
        # CA Policies sheet
        if ($TenantData.ConditionalAccess.Policies.Count -gt 0) {
            $TenantData.ConditionalAccess.Policies | Export-Excel -Path $excelPath -WorksheetName "CA Policies" -AutoSize -BoldTopRow
        }
        
        # Licenses sheet
        if ($TenantData.Licenses.SubscribedSkus.Count -gt 0) {
            $TenantData.Licenses.SubscribedSkus | Export-Excel -Path $excelPath -WorksheetName "Licenses" -AutoSize -BoldTopRow
        }
        
        Write-LogMessage -Message "Basic Excel documentation generated: $excelPath" -Type Success -LogOnly
        return $true
    }
    catch {
        Write-LogMessage -Message "Failed to generate basic Excel documentation: $($_.Exception.Message)" -Type Error
        return $false
    }
}

function New-JSONDocumentation {
    <#
    .SYNOPSIS
        Generates JSON configuration export
    #>
    param (
        [Parameter(Mandatory = $true)]
        [hashtable]$TenantData
    )
    
    try {
        $jsonPath = Join-Path -Path $DocumentationConfig.OutputDirectory -ChildPath "Exports\TenantConfiguration.json"
        
        $TenantData | ConvertTo-Json -Depth 10 | Out-File -FilePath $jsonPath -Encoding UTF8
        Write-LogMessage -Message "JSON configuration export generated: $jsonPath" -Type Success -LogOnly
        return $true
    }
    catch {
        Write-LogMessage -Message "Failed to generate JSON documentation: $($_.Exception.Message)" -Type Error
        return $false
    }
}

function New-ConfigurationSummary {
    <#
    .SYNOPSIS
        Generates a text-based configuration summary
    #>
    param (
        [Parameter(Mandatory = $true)]
        [hashtable]$TenantData
    )
    
    try {
        $summaryPath = Join-Path -Path $DocumentationConfig.OutputDirectory -ChildPath "Reports\ConfigurationSummary.txt"
        
        $summary = @"
Microsoft 365 Tenant Configuration Summary
==========================================
Generated: $($TenantData.GeneratedOn.ToString('yyyy-MM-dd HH:mm:ss'))

TENANT INFORMATION
------------------
Tenant Name: $($TenantData.TenantInfo.DisplayName)
Default Domain: $($TenantData.TenantInfo.DefaultDomain)
Tenant ID: $($TenantData.TenantInfo.TenantId)
Connected As: $($TenantData.TenantInfo.ConnectedAs)
Setup Date: $($TenantData.TenantInfo.SetupDate)

CONFIGURATION SUMMARY
---------------------
Total Groups: $($TenantData.Groups.TotalCount)
  - Security Groups: $($TenantData.Groups.SecurityGroups.Count)
  - Microsoft 365 Groups: $($TenantData.Groups.Microsoft365Groups.Count)
  - Distribution Groups: $($TenantData.Groups.DistributionGroups.Count)
  - Dynamic Groups: $($TenantData.Groups.DynamicGroups.Count)

Users: $($TenantData.Users.TotalUsers) total
  - Enabled: $($TenantData.Users.EnabledUsers)
  - Disabled: $($TenantData.Users.DisabledUsers)
  - Guests: $($TenantData.Users.GuestUsers)

Conditional Access Policies: $($TenantData.ConditionalAccess.TotalCount) total
  - Enabled: $($TenantData.ConditionalAccess.EnabledCount)
  - Disabled: $($TenantData.ConditionalAccess.DisabledCount)

Intune Managed Devices: $($TenantData.Intune.TotalDevices)

License Usage:
  - Total Licensed: $($TenantData.Licenses.UsedLicenses)
  - Total Available: $($TenantData.Licenses.TotalLicenses)

SECURITY GROUPS CREATED
-----------------------
"@

        foreach ($group in $TenantData.Groups.SecurityGroups) {
            $summary += "`n- $($group.DisplayName)"
            if (Test-NotEmpty -Value $group.Description) {
                $summary += " ($($group.Description))"
            }
        }

        $summary += @"


CONDITIONAL ACCESS POLICIES
----------------------------
"@

        foreach ($policy in $TenantData.ConditionalAccess.Policies) {
            $summary += "`n- $($policy.DisplayName) [$($policy.State)]"
        }

        $summary += @"


NEXT STEPS
----------
1. Review all configurations in the detailed reports
2. Test conditional access policies with pilot users
3. Configure additional SharePoint sites as needed
4. Set up Intune app protection policies
5. Configure user onboarding processes
6. Set up monitoring and alerting
7. Schedule regular configuration reviews

For detailed information, see the HTML and Excel reports in the Reports folder.
"@

        $summary | Out-File -FilePath $summaryPath -Encoding UTF8
        Write-LogMessage -Message "Configuration summary generated: $summaryPath" -Type Success -LogOnly
        return $true
    }
    catch {
        Write-LogMessage -Message "Failed to generate configuration summary: $($_.Exception.Message)" -Type Error
        return $false
    }
}

function New-SetupChecklist {
    <#
    .SYNOPSIS
        Generates a setup completion checklist
    #>
    param (
        [Parameter(Mandatory = $true)]
        [hashtable]$TenantData
    )
    
    try {
        $checklistPath = Join-Path -Path $DocumentationConfig.OutputDirectory -ChildPath "Templates\PostSetupChecklist.txt"
        
        $checklist = @"
Microsoft 365 Tenant Post-Setup Checklist
==========================================
Tenant: $($TenantData.TenantInfo.DisplayName)
Date: $(Get-Date -Format 'yyyy-MM-dd')

CONFIGURATION VERIFICATION
--------------------------
☐ Verify all security groups are created and populated correctly
☐ Test conditional access policies with pilot users
☐ Confirm license assignments are working correctly
☐ Verify SharePoint site permissions and access
☐ Test Intune device enrollment and compliance
☐ Confirm user creation and onboarding process
☐ Verify email flow and Exchange configuration
☐ Test external sharing settings (if applicable)

SECURITY REVIEW
---------------
☐ Review and enable security defaults (if not using CA policies)
☐ Configure multi-factor authentication requirements
☐ Set up privileged identity management
☐ Review admin role assignments
☐ Configure audit logging and retention
☐ Set up security alerts and monitoring
☐ Review data loss prevention policies
☐ Configure information protection labels

OPERATIONAL TASKS
-----------------
☐ Set up user onboarding documentation
☐ Create helpdesk procedures for common issues
☐ Configure backup and retention policies
☐ Set up monitoring and health checks
☐ Create change management procedures
☐ Document emergency access procedures
☐ Schedule regular access reviews
☐ Plan for future growth and scaling

TRAINING AND ADOPTION
---------------------
☐ Plan admin training sessions
☐ Create end-user training materials
☐ Set up Teams adoption strategy
☐ Plan SharePoint governance training
☐ Create security awareness training
☐ Document support procedures
☐ Plan rollout communication

COMPLIANCE AND GOVERNANCE
-------------------------
☐ Review data residency requirements
☐ Configure compliance policies
☐ Set up retention labels and policies
☐ Review GDPR/privacy compliance
☐ Configure eDiscovery settings
☐ Set up legal hold procedures
☐ Document compliance processes

MONITORING AND MAINTENANCE
--------------------------
☐ Set up service health monitoring
☐ Configure usage analytics
☐ Plan regular configuration reviews
☐ Set up automated reporting
☐ Configure capacity planning alerts
☐ Plan for regular updates and patches
☐ Schedule periodic security reviews

NOTES
-----
(Add any specific notes or requirements for this tenant)

_________________________________________________
_________________________________________________
_________________________________________________
_________________________________________________

Completed by: _________________________ Date: _____________
Reviewed by: __________________________ Date: _____________
"@

        $checklist | Out-File -FilePath $checklistPath -Encoding UTF8
        Write-LogMessage -Message "Setup checklist generated: $checklistPath" -Type Success -LogOnly
        return $true
    }
    catch {
        Write-LogMessage -Message "Failed to generate setup checklist: $($_.Exception.Message)" -Type Error
        return $false
    }
}

# === Export Functions ===
# Make sure the main function is available for calling
Export-ModuleMember -Function New-TenantDocumentation