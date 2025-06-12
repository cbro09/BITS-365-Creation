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
        Close-ExcelPackage -ExcelPackage $excel -SaveAs $outputPath
        
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
            if ($currentRow -gt ($startRow + 500)) {
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
            # Always add user even if no licenses
            $worksheet.Cells[$currentRow, 2].Value = $user.DisplayName         # Column B: User Name
            
            # Handle license assignments properly
            if ($user.AssignedLicenses -and $user.AssignedLicenses.Count -gt 0) {
                # Convert license array to proper format
                $primaryLicense = $user.AssignedLicenses[0]
                if ($primaryLicense -and $primaryLicense -ne 0) {
                    $worksheet.Cells[$currentRow, 3].Value = $primaryLicense # Column C: Base License Type
                } else {
                    $worksheet.Cells[$currentRow, 3].Value = "No License Assigned"
                }
                
                # Additional licenses in subsequent columns
                if ($user.AssignedLicenses.Count -gt 1) {
                    $secondLicense = $user.AssignedLicenses[1]
                    if ($secondLicense -and $secondLicense -ne 0) {
                        $worksheet.Cells[$currentRow, 4].Value = $secondLicense  # Column D: Additional Software 1
                    }
                }
                if ($user.AssignedLicenses.Count -gt 2) {
                    $thirdLicense = $user.AssignedLicenses[2]
                    if ($thirdLicense -and $thirdLicense -ne 0) {
                        $worksheet.Cells[$currentRow, 5].Value = $thirdLicense  # Column E: Additional Software 2
                    }
                }
            } else {
                $worksheet.Cells[$currentRow, 3].Value = "No License Assigned"  # Column C: Base License Type
            }
            
            $currentRow++
            
            # Limit to prevent performance issues
            if ($currentRow -gt ($startRow + 500)) {
                break
            }
        }
        
        Write-LogMessage -Message "Updated Licensing sheet with license assignments for $($currentRow - $startRow) users" -Type Success -LogOnly
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
                if ($configRow -gt 20) { break }
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
                if ($complianceRow -gt 30) { break }
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
            if ($currentRow -gt ($startRow + 20)) { break }
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
        
        # Add actual SharePoint sites
        if ($TenantData.SharePoint.SiteCollections -and $TenantData.SharePoint.SiteCollections.Count -gt 0) {
            foreach ($site in $TenantData.SharePoint.SiteCollections) {
                $worksheet.Cells[$currentRow, 2].Value = $site.DisplayName  # Column B: Site Name
                $worksheet.Cells[$currentRow, 4].Value = "Site Admin"       # Column D: Approver
                $worksheet.Cells[$currentRow, 6].Value = "Site Owners"      # Column F: Owners
                $worksheet.Cells[$currentRow, 8].Value = "Site Members"     # Column H: Members
                $currentRow++
                
                # Limit entries
                if ($currentRow -gt ($startRow + 20)) { break }
            }
            Write-LogMessage -Message "Updated SharePoint Libraries sheet with $($currentRow - $startRow) sites" -Type Success -LogOnly
        } else {
            # If no sites found, add a note
            $worksheet.Cells[$startRow, 2].Value = "No SharePoint sites found or unable to access"
            $worksheet.Cells[$startRow, 4].Value = "N/A"
            Write-LogMessage -Message "No SharePoint sites found to populate" -Type Warning -LogOnly
        }
    }
    catch {
        Write-LogMessage -Message "Error updating SharePoint Libraries sheet: $($_.Exception.Message)" -Type Warning -LogOnly
    }
}

function Update-IntuneAppsSheets {
    <#
    .SYNOPSIS
        Populates all Intune Apps sheets with actual app data
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
                $startRow = 8  # Start after headers (Application Name, Required, Optional, Selected users only)
                $currentRow = $startRow
                
                # Add actual managed apps
                if ($TenantData.Intune.ManagedApps -and $TenantData.Intune.ManagedApps.Count -gt 0) {
                    foreach ($app in $TenantData.Intune.ManagedApps) {
                        # Filter apps by platform if needed
                        $platformMatch = $true
                        if ($sheetName -like "*Windows*" -and $app.DisplayName -notlike "*Windows*" -and $app.DisplayName -notlike "*Office*" -and $app.DisplayName -notlike "*Microsoft*") {
                            $platformMatch = $false
                        }
                        
                        if ($platformMatch) {
                            $worksheet.Cells[$currentRow, 2].Value = $app.DisplayName    # Column B: Application Name
                            $worksheet.Cells[$currentRow, 3].Value = "X"                # Column C: Required (assuming required)
                            $currentRow++
                            
                            # Limit entries per sheet
                            if ($currentRow -gt ($startRow + 15)) { break }
                        }
                    }
                    Write-LogMessage -Message "Updated $sheetName sheet with managed apps" -Type Success -LogOnly
                } else {
                    # If no apps found, add a note
                    $worksheet.Cells[$startRow, 2].Value = "No managed apps found"
                    $worksheet.Cells[$startRow, 3].Value = ""
                    Write-LogMessage -Message "No managed apps found for $sheetName" -Type Warning -LogOnly
                }
            }
        }
    }
    catch {
        Write-LogMessage -Message "Error updating Intune Apps sheets: $($_.Exception.Message)" -Type Warning -LogOnly
    }
}

function Update-DistributionListsSheet {
    <#
    .SYNOPSIS
        Populates the Distribution Lists sheet with proper table formatting
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
        
        # Add proper headers if missing
        $worksheet.Cells[7, 2].Value = "Distribution List Name"    # Column B: Group Name
        $worksheet.Cells[7, 3].Value = "Description"              # Column C: Description  
        $worksheet.Cells[7, 4].Value = "Member Count"             # Column D: Member Count
        $worksheet.Cells[7, 6].Value = "Members"                  # Column F: Members
        
        $startRow = 8
        $currentRow = $startRow
        
        foreach ($group in $TenantData.Groups.DistributionGroups) {
            $worksheet.Cells[$currentRow, 2].Value = $group.DisplayName        # Column B: Group Name
            $worksheet.Cells[$currentRow, 3].Value = Get-SafeString -Value $group.Description -MaxLength 100  # Column C: Description
            $worksheet.Cells[$currentRow, 4].Value = $group.MemberCount        # Column D: Member Count
            $worksheet.Cells[$currentRow, 6].Value = "See member details in Groups section"  # Column F: Members reference
            $currentRow++
            
            if ($currentRow -gt ($startRow + 20)) { break }
        }
        
        Write-LogMessage -Message "Updated Distribution Lists sheet with $($currentRow - $startRow) groups" -Type Success -LogOnly
    }
    catch {
        Write-LogMessage -Message "Error updating Distribution Lists sheet: $($_.Exception.Message)" -Type Warning -LogOnly
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
        
        # Look for actual shared mailboxes (mail-enabled groups or specific types)
        $sharedMailboxes = @()
        
        # Check distribution groups that might be shared mailboxes
        if ($TenantData.Groups.DistributionGroups) {
            $sharedMailboxes += $TenantData.Groups.DistributionGroups | Where-Object { 
                $_.DisplayName -like "*shared*" -or 
                $_.DisplayName -like "*mailbox*" -or
                $_.DisplayName -like "*info*" -or
                $_.DisplayName -like "*support*"
            }
        }
        
        # Check Microsoft 365 groups that might be shared mailboxes
        if ($TenantData.Groups.Microsoft365Groups) {
            $sharedMailboxes += $TenantData.Groups.Microsoft365Groups | Where-Object { 
                $_.DisplayName -like "*shared*" -or 
                $_.DisplayName -like "*mailbox*"
            }
        }
        
        if ($sharedMailboxes.Count -gt 0) {
            foreach ($mailbox in $sharedMailboxes) {
                $worksheet.Cells[$currentRow, 2].Value = $mailbox.DisplayName
                $worksheet.Cells[$currentRow, 3].Value = Get-SafeString -Value $mailbox.Description -MaxLength 100
                $currentRow++
                
                if ($currentRow -gt ($startRow + 10)) { break }
            }
            Write-LogMessage -Message "Updated Shared Mailboxes sheet with $($currentRow - $startRow) mailboxes" -Type Success -LogOnly
        } else {
            $worksheet.Cells[$startRow, 2].Value = "No shared mailboxes found"
            Write-LogMessage -Message "No shared mailboxes found" -Type Warning -LogOnly
        }
    }
    catch {
        Write-LogMessage -Message "Error updating Shared Mailboxes sheet: $($_.Exception.Message)" -Type Warning -LogOnly
    }
}

# === Data Collection Functions ===

function Get-CompleteTenantConfiguration {
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
        $spInfo = @{
            TenantSettings = @{}
            SiteCollections = @()
            TotalSites = 0
            StorageUsed = "Not available"
            SharingSettings = "Not available"
            ExternalSharingEnabled = "Not available"
        }
        
        # Try to get SharePoint sites through Graph API
        try {
            Write-LogMessage -Message "Collecting SharePoint sites..." -Type Info -LogOnly
            
            # Try different approaches to get SharePoint sites
            $sites = @()
            
            # Method 1: Try to get all sites
            try {
                $allSites = Get-MgSite -All -Top 100
                $sites += $allSites
                Write-LogMessage -Message "Found $($allSites.Count) sites using Get-MgSite -All" -Type Info -LogOnly
            }
            catch {
                Write-LogMessage -Message "Get-MgSite -All failed: $($_.Exception.Message)" -Type Warning -LogOnly
            }
            
            # Method 2: Try to search for sites
            if ($sites.Count -eq 0) {
                try {
                    $searchSites = Get-MgSite -Search "*"
                    $sites += $searchSites
                    Write-LogMessage -Message "Found $($searchSites.Count) sites using search" -Type Info -LogOnly
                }
                catch {
                    Write-LogMessage -Message "Site search failed: $($_.Exception.Message)" -Type Warning -LogOnly
                }
            }
            
            # Method 3: Try to get root site and subsites
            if ($sites.Count -eq 0) {
                try {
                    $rootSite = Get-MgSite -SiteId "root"
                    if ($rootSite) {
                        $sites += $rootSite
                        Write-LogMessage -Message "Found root site" -Type Info -LogOnly
                        
                        # Try to get subsites
                        try {
                            $subSites = Get-MgSiteSite -SiteId $rootSite.Id
                            $sites += $subSites
                            Write-LogMessage -Message "Found $($subSites.Count) subsites" -Type Info -LogOnly
                        }
                        catch {
                            Write-LogMessage -Message "Could not get subsites: $($_.Exception.Message)" -Type Warning -LogOnly
                        }
                    }
                }
                catch {
                    Write-LogMessage -Message "Could not get root site: $($_.Exception.Message)" -Type Warning -LogOnly
                }
            }
            
            # Process found sites
            if ($sites.Count -gt 0) {
                $spInfo.TotalSites = $sites.Count
                $spInfo.SiteCollections = $sites | ForEach-Object {
                    @{
                        Id = $_.Id
                        DisplayName = $_.DisplayName
                        Name = $_.Name
                        WebUrl = $_.WebUrl
                        CreatedDateTime = $_.CreatedDateTime
                        LastModifiedDateTime = $_.LastModifiedDateTime
                        SiteCollection = $_.SiteCollection
                    }
                }
                Write-LogMessage -Message "Successfully collected $($sites.Count) SharePoint sites" -Type Success -LogOnly
            } else {
                Write-LogMessage -Message "No SharePoint sites found through any method" -Type Warning -LogOnly
                # Add a placeholder entry to indicate we tried but found nothing
                $spInfo.SiteCollections = @(
                    @{
                        Id = "N/A"
                        DisplayName = "No sites accessible"
                        Name = "N/A"
                        WebUrl = "N/A"
                        CreatedDateTime = Get-Date
                        LastModifiedDateTime = Get-Date
                        SiteCollection = $null
                    }
                )
            }
        }
        catch {
            Write-LogMessage -Message "SharePoint sites collection failed: $($_.Exception.Message)" -Type Warning -LogOnly
            # Add error indicator
            $spInfo.SiteCollections = @(
                @{
                    Id = "ERROR"
                    DisplayName = "Error accessing SharePoint data"
                    Name = "ERROR"
                    WebUrl = "Check permissions"
                    CreatedDateTime = Get-Date
                    LastModifiedDateTime = Get-Date
                    SiteCollection = $null
                }
            )
        }
        
        return $spInfo
    }
    catch {
        Write-LogMessage -Message "Error collecting SharePoint information: $($_.Exception.Message)" -Type Warning -LogOnly
        return @{
            TenantSettings = @{}
            SiteCollections = @()
            TotalSites = 0
            StorageUsed = "Error"
            SharingSettings = "Error"
            ExternalSharingEnabled = "Error"
        }
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
            ManagedApps = @()
            TotalDevices = 0
        }
        
        # Device Compliance Policies - Get ALL, not limited
        try {
            Write-LogMessage -Message "Collecting device compliance policies..." -Type Info -LogOnly
            $compliancePolicies = Get-MgDeviceManagementDeviceCompliancePolicy -All
            Write-LogMessage -Message "Found $($compliancePolicies.Count) compliance policies" -Type Info -LogOnly
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
            Write-LogMessage -Message "Could not retrieve device compliance policies: $($_.Exception.Message)" -Type Warning -LogOnly
        }
        
        # Device Configuration Policies - Get ALL, not limited
        try {
            Write-LogMessage -Message "Collecting device configuration policies..." -Type Info -LogOnly
            $configPolicies = Get-MgDeviceManagementDeviceConfiguration -All
            Write-LogMessage -Message "Found $($configPolicies.Count) configuration policies" -Type Info -LogOnly
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
            Write-LogMessage -Message "Could not retrieve device configuration policies: $($_.Exception.Message)" -Type Warning -LogOnly
        }
        
        # Managed Apps - NEW: Collect managed applications
        try {
            Write-LogMessage -Message "Collecting managed applications..." -Type Info -LogOnly
            $managedApps = Get-MgDeviceManagementMobileApp -All
            Write-LogMessage -Message "Found $($managedApps.Count) managed apps" -Type Info -LogOnly
            $intuneInfo.ManagedApps = $managedApps | ForEach-Object {
                @{
                    Id = $_.Id
                    DisplayName = $_.DisplayName
                    Description = $_.Description
                    Publisher = $_.Publisher
                    CreatedDateTime = $_.CreatedDateTime
                    LastModifiedDateTime = $_.LastModifiedDateTime
                }
            }
        }
        catch {
            Write-LogMessage -Message "Could not retrieve managed applications: $($_.Exception.Message)" -Type Warning -LogOnly
        }
        
        # Managed Devices
        try {
            Write-LogMessage -Message "Collecting managed devices..." -Type Info -LogOnly
            $devices = Get-MgDeviceManagementManagedDevice -All -Top 500
            $intuneInfo.TotalDevices = $devices.Count
            Write-LogMessage -Message "Found $($devices.Count) managed devices" -Type Info -LogOnly
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
            Write-LogMessage -Message "Could not retrieve managed devices: $($_.Exception.Message)" -Type Warning -LogOnly
        }
        
        Write-LogMessage -Message "Intune data collection completed - Config: $($intuneInfo.DeviceConfigurationPolicies.Count), Compliance: $($intuneInfo.DeviceCompliancePolicies.Count), Apps: $($intuneInfo.ManagedApps.Count)" -Type Info -LogOnly
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