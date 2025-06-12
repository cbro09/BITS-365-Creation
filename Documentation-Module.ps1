# === Documentation.ps1 ===
# Documentation generation and reporting functions - Complete Implementation

function New-TenantDocumentation {
    Write-LogMessage -Message "Starting documentation generation..." -Type Info
    
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
            Disconnect-SPOService -ErrorAction SilentlyContinue | Out-Null
        }
        catch {
            # Ignore disconnect errors
        }
        
        # STEP 5: Force load ONLY the exact modules needed for Documentation
        $documentationModules = @(
            'Microsoft.Graph.Identity.DirectoryManagement',
            'Microsoft.Graph.Users',
            'Microsoft.Graph.Groups',
            'ImportExcel'
        )
        
        Write-LogMessage -Message "Loading ONLY Documentation modules in exact order..." -Type Info
        foreach ($module in $documentationModules) {
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
        
        # STEP 6: Connect with EXACT scopes needed for Documentation
        $documentationScopes = @(
            "Directory.Read.All",
            "User.Read.All", 
            "Group.Read.All",
            "Policy.Read.ConditionalAccess"
        )
        
        Write-LogMessage -Message "Connecting to Microsoft Graph with Documentation scopes..." -Type Info
        Connect-MgGraph -Scopes $documentationScopes -ErrorAction Stop | Out-Null
        
        $context = Get-MgContext
        Write-LogMessage -Message "Connected to Microsoft Graph as $($context.Account)" -Type Success
        
        # STEP 7: Documentation generation logic
        
        # Validate tenant state exists
        if (-not $script:TenantState) {
            Write-LogMessage -Message "No tenant state found. Please run other modules first to configure the tenant." -Type Error
            return $false
        }
        
        # Get source template path
        Write-LogMessage -Message "Locating master spreadsheet template..." -Type Info
        $templatePath = $null
        
        # Check common locations for the template
        $possiblePaths = @(
            "Master Spreadsheet Customer Details.xlsx",
            "$PSScriptRoot\Master Spreadsheet Customer Details.xlsx",
            "$env:USERPROFILE\Documents\Master Spreadsheet Customer Details.xlsx",
            "$env:USERPROFILE\Downloads\Master Spreadsheet Customer Details.xlsx"
        )
        
        foreach ($path in $possiblePaths) {
            if (Test-Path $path) {
                $templatePath = $path
                break
            }
        }
        
        if (-not $templatePath) {
            # Prompt user for template location
            Add-Type -AssemblyName System.Windows.Forms
            $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
            $openFileDialog.Title = "Select Master Spreadsheet Template"
            $openFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
            $openFileDialog.InitialDirectory = "$env:USERPROFILE\Documents"
            
            if ($openFileDialog.ShowDialog() -eq 'OK') {
                $templatePath = $openFileDialog.FileName
            } else {
                Write-LogMessage -Message "No template file selected. Documentation generation cancelled." -Type Warning
                return $false
            }
        }
        
        Write-LogMessage -Message "Using template: $templatePath" -Type Success
        
        # Create output path
        $customerName = $script:TenantState.TenantName -replace '[^\w\s-]', ''
        $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
        $outputPath = "$env:USERPROFILE\Documents\${customerName}_TenantDocumentation_${timestamp}.xlsx"
        
        Write-LogMessage -Message "Generating documentation for: $($script:TenantState.TenantName)" -Type Info
        Write-LogMessage -Message "Output will be saved to: $outputPath" -Type Info
        
        # Copy template to output location
        Copy-Item -Path $templatePath -Destination $outputPath -Force
        
        # Gather current tenant data
        Write-LogMessage -Message "Gathering current tenant configuration data..." -Type Info
        $tenantData = Get-TenantConfigurationData
        
        # Update the Excel file with current data
        Write-LogMessage -Message "Populating Excel template with tenant data..." -Type Info
        Update-ExcelWithTenantData -ExcelPath $outputPath -TenantData $tenantData
        
        Write-LogMessage -Message "Documentation generation completed successfully" -Type Success
        Write-LogMessage -Message "Customer documentation saved to: $outputPath" -Type Success
        
        # Open the file for review
        $openFile = Read-Host "Would you like to open the generated documentation? (Y/N)"
        if ($openFile -eq 'Y' -or $openFile -eq 'y') {
            Start-Process $outputPath
        }
        
        return $true
    }
    catch {
        Write-LogMessage -Message "Error in documentation generation - $($_.Exception.Message)" -Type Error
        return $false
    }
}

function Get-TenantConfigurationData {
    Write-LogMessage -Message "Collecting tenant configuration data..." -Type Info
    
    $data = @{
        TenantInfo = @{
            TenantName = $script:TenantState.TenantName
            DefaultDomain = $script:TenantState.DefaultDomain
            TenantId = $script:TenantState.TenantId
            AdminEmail = $script:TenantState.AdminEmail
            GeneratedDate = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        }
        Users = @()
        Groups = @()
        ConditionalAccessPolicies = @()
        SharePointSites = @()
        Licensing = @()
    }
    
    try {
        # Get all users
        Write-LogMessage -Message "Retrieving user information..." -Type Info
        $users = Get-MgUser -All -Property DisplayName,UserPrincipalName,GivenName,Surname,JobTitle,Department,OfficeLocation,MobilePhone,Manager,OnPremisesExtensionAttributes
        
        foreach ($user in $users) {
            # Get manager information if exists
            $managerEmail = ""
            if ($user.Manager) {
                try {
                    $manager = Get-MgUser -UserId $user.Manager.Id -Property UserPrincipalName -ErrorAction SilentlyContinue
                    $managerEmail = $manager.UserPrincipalName
                } catch {
                    # Manager lookup failed, continue without
                }
            }
            
            # Extract license from extension attributes
            $licenseType = ""
            if ($user.OnPremisesExtensionAttributes -and $user.OnPremisesExtensionAttributes.ExtensionAttribute1) {
                $licenseType = $user.OnPremisesExtensionAttributes.ExtensionAttribute1
            }
            
            $data.Users += @{
                FirstName = $user.GivenName
                LastName = $user.Surname
                Email = $user.UserPrincipalName
                JobTitle = $user.JobTitle
                ManagerEmail = $managerEmail
                Department = $user.Department
                OfficeLocation = $user.OfficeLocation
                PhoneNumber = $user.MobilePhone
                LicenseType = $licenseType
            }
        }
        
        # Get all groups created by our script
        Write-LogMessage -Message "Retrieving group information..." -Type Info
        if ($script:TenantState.CreatedGroups) {
            foreach ($groupName in $script:TenantState.CreatedGroups.Keys) {
                try {
                    $groupId = $script:TenantState.CreatedGroups[$groupName]
                    $group = Get-MgGroup -GroupId $groupId -Property DisplayName,Description,GroupTypes
                    
                    $data.Groups += @{
                        Name = $group.DisplayName
                        Description = $group.Description
                        Type = if ($group.GroupTypes -contains "DynamicMembership") { "Dynamic" } else { "Static" }
                    }
                } catch {
                    Write-LogMessage -Message "Could not retrieve group: $groupName" -Type Warning -LogOnly
                }
            }
        }
        
        # Get Conditional Access policies (common names from our script)
        Write-LogMessage -Message "Retrieving Conditional Access policies..." -Type Info
        $commonPolicyNames = @(
            "C001 - Block Legacy Authentication All Apps",
            "C002 - MFA Required for All Users", 
            "C003 - Block Non Corporate Devices",
            "C004 - Require Password Change and MFA for High Risk Users",
            "C005 - Require MFA for Risky Sign-Ins"
        )
        
        try {
            $allPolicies = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies"
            
            foreach ($policyName in $commonPolicyNames) {
                $policy = $allPolicies.value | Where-Object { $_.displayName -eq $policyName }
                
                if ($policy) {
                    $policyDescription = Get-PolicyDescription -PolicyName $policyName
                    
                    $data.ConditionalAccessPolicies += @{
                        Name = $policy.displayName
                        Description = $policyDescription
                        State = $policy.state
                    }
                }
            }
        } catch {
            Write-LogMessage -Message "Could not retrieve Conditional Access policies - may not have sufficient permissions" -Type Warning
        }
        
        # SharePoint sites information (from known structure)
        $defaultSites = @("HR", "Policies", "Templates", "Processes", "Documents")
        foreach ($siteName in $defaultSites) {
            $data.SharePointSites += @{
                Name = $siteName
                URL = "https://{tenant}.sharepoint.com/sites/$($siteName.ToLower())"
                Description = "Standard $siteName library for organization documents"
            }
        }
        
        # Add hub site if tenant name available
        if ($script:TenantState.TenantName) {
            $data.SharePointSites += @{
                Name = "$($script:TenantState.TenantName) Hub"
                URL = "https://{tenant}.sharepoint.com/sites/corporatehub"
                Description = "Central hub site for navigation and organization"
            }
        }
        
        Write-LogMessage -Message "Tenant data collection completed" -Type Success
        Write-LogMessage -Message "Collected: $($data.Users.Count) users, $($data.Groups.Count) groups, $($data.ConditionalAccessPolicies.Count) CA policies, $($data.SharePointSites.Count) SharePoint sites" -Type Info
        
    } catch {
        Write-LogMessage -Message "Error collecting tenant data: $($_.Exception.Message)" -Type Warning
    }
    
    return $data
}

function Update-ExcelWithTenantData {
    param(
        [string]$ExcelPath,
        [hashtable]$TenantData
    )
    
    Write-LogMessage -Message "Updating Excel file with tenant configuration data..." -Type Info
    
    try {
        # Open the Excel file
        $excel = Open-ExcelPackage -Path $ExcelPath
        
        # Update M365 Due Diligence sheet - Replace company name placeholder
        if ($excel.Workbook.Worksheets["M365 Due Dilligence"]) {
            $sheet = $excel.Workbook.Worksheets["M365 Due Dilligence"]
            $cell = $sheet.Cells["C7"]
            if ($cell.Value -and $cell.Value.ToString().Contains("{CompanyName}")) {
                $cell.Value = $cell.Value.ToString().Replace("{CompanyName}", $TenantData.TenantInfo.TenantName)
                Write-LogMessage -Message "Updated company name in Due Diligence sheet" -Type Success -LogOnly
            }
        }
        
        # Update Users sheet
        if ($excel.Workbook.Worksheets["Users"] -and $TenantData.Users.Count -gt 0) {
            $sheet = $excel.Workbook.Worksheets["Users"]
            $startRow = 7  # Data starts at row 7 based on analysis
            
            for ($i = 0; $i -lt $TenantData.Users.Count; $i++) {
                $user = $TenantData.Users[$i]
                $row = $startRow + $i
                
                $sheet.Cells["A$row"].Value = $user.FirstName
                $sheet.Cells["B$row"].Value = $user.LastName  
                $sheet.Cells["C$row"].Value = $user.Email
                $sheet.Cells["D$row"].Value = $user.JobTitle
                $sheet.Cells["E$row"].Value = $user.ManagerEmail
                $sheet.Cells["F$row"].Value = $user.Department
                $sheet.Cells["G$row"].Value = $user.OfficeLocation
                $sheet.Cells["H$row"].Value = $user.PhoneNumber
            }
            Write-LogMessage -Message "Updated Users sheet with $($TenantData.Users.Count) users" -Type Success -LogOnly
        }
        
        # Update Licensing sheet
        if ($excel.Workbook.Worksheets["Licensing"] -and $TenantData.Users.Count -gt 0) {
            $sheet = $excel.Workbook.Worksheets["Licensing"]
            $startRow = 8  # Data starts at row 8 based on analysis
            
            for ($i = 0; $i -lt $TenantData.Users.Count; $i++) {
                $user = $TenantData.Users[$i]
                $row = $startRow + $i
                
                $sheet.Cells["B$row"].Value = $user.Email  # User Name
                $sheet.Cells["D$row"].Value = $user.LicenseType  # Base License Type
            }
            Write-LogMessage -Message "Updated Licensing sheet with user license information" -Type Success -LogOnly
        }
        
        # Update SharePoint Libraries sheet
        if ($excel.Workbook.Worksheets["SharePoint Libaries"] -and $TenantData.SharePointSites.Count -gt 0) {
            $sheet = $excel.Workbook.Worksheets["SharePoint Libaries"]
            # Libraries are pre-populated in rows 7-12, just confirm they exist
            Write-LogMessage -Message "SharePoint Libraries sheet maintained with default library structure" -Type Success -LogOnly
        }
        
        # Update Conditional Access sheet
        if ($excel.Workbook.Worksheets["Conditional Access"] -and $TenantData.ConditionalAccessPolicies.Count -gt 0) {
            $sheet = $excel.Workbook.Worksheets["Conditional Access"]
            # Policies are pre-populated in rows 9-12, update status if needed
            $policyRows = @{
                "User Risk" = 9
                "Block Legacy Auth" = 10
                "MFA for all Users" = 11
                "Non Corporate Device Block" = 12
            }
            
            foreach ($policy in $TenantData.ConditionalAccessPolicies) {
                $mappedName = Get-PolicyMappedName -PolicyName $policy.Name
                if ($policyRows.ContainsKey($mappedName)) {
                    $row = $policyRows[$mappedName]
                    $currentSetting = $sheet.Cells["D$row"].Value
                    $sheet.Cells["D$row"].Value = "$currentSetting (Status: $($policy.State))"
                }
            }
            Write-LogMessage -Message "Updated Conditional Access sheet with policy status" -Type Success -LogOnly
        }
        
        # Add summary information to a new sheet or existing summary area
        Add-TenantSummaryInfo -Excel $excel -TenantData $TenantData
        
        # Save the updated Excel file
        Close-ExcelPackage $excel -Save
        Write-LogMessage -Message "Excel file updated and saved successfully" -Type Success
        
    } catch {
        Write-LogMessage -Message "Error updating Excel file: $($_.Exception.Message)" -Type Error
        if ($excel) {
            Close-ExcelPackage $excel -NoSave
        }
        throw
    }
}

function Add-TenantSummaryInfo {
    param(
        [OfficeOpenXml.ExcelPackage]$Excel,
        [hashtable]$TenantData
    )
    
    try {
        # Add summary to the first sheet or create a summary section
        $sheet = $Excel.Workbook.Worksheets["M365 Due Dilligence"]
        
        if ($sheet) {
            # Find empty area to add summary (around row 35-40)
            $summaryRow = 35
            
            $sheet.Cells["B$summaryRow"].Value = "TENANT CONFIGURATION SUMMARY"
            $sheet.Cells["B$summaryRow"].Style.Font.Bold = $true
            $summaryRow++
            
            $sheet.Cells["B$summaryRow"].Value = "Generated: $($TenantData.TenantInfo.GeneratedDate)"
            $summaryRow++
            $sheet.Cells["B$summaryRow"].Value = "Tenant Name: $($TenantData.TenantInfo.TenantName)"
            $summaryRow++
            $sheet.Cells["B$summaryRow"].Value = "Default Domain: $($TenantData.TenantInfo.DefaultDomain)"
            $summaryRow++
            $sheet.Cells["B$summaryRow"].Value = "Admin Email: $($TenantData.TenantInfo.AdminEmail)"
            $summaryRow++
            $summaryRow++
            
            $sheet.Cells["B$summaryRow"].Value = "CONFIGURATION COUNTS:"
            $sheet.Cells["B$summaryRow"].Style.Font.Bold = $true
            $summaryRow++
            $sheet.Cells["B$summaryRow"].Value = "Users Created: $($TenantData.Users.Count)"
            $summaryRow++
            $sheet.Cells["B$summaryRow"].Value = "Groups Created: $($TenantData.Groups.Count)"
            $summaryRow++
            $sheet.Cells["B$summaryRow"].Value = "CA Policies: $($TenantData.ConditionalAccessPolicies.Count)"
            $summaryRow++
            $sheet.Cells["B$summaryRow"].Value = "SharePoint Sites: $($TenantData.SharePointSites.Count)"
            
            Write-LogMessage -Message "Added tenant summary information to documentation" -Type Success -LogOnly
        }
    } catch {
        Write-LogMessage -Message "Error adding summary information: $($_.Exception.Message)" -Type Warning -LogOnly
    }
}

function Get-PolicyDescription {
    param([string]$PolicyName)
    
    $descriptions = @{
        "C001 - Block Legacy Authentication All Apps" = "Scope: All Users, Target Resources: None, Network: None, Conditions: Client Apps - Other Clients, Access Control: Block Access"
        "C002 - MFA Required for All Users" = "Scope: All Users, Target Resources: All resources, Network: None, Conditions: None, Access Control: Grant Access when 'Require multifactor authentication'"
        "C003 - Block Non Corporate Devices" = "Scope: All Users, Target Resources: Office 365, Microsoft Teams Services, Network: None, Conditions: Device Platforms - Any Device, Grant Access if 'Require device to be marked compliant' & 'Require Microsoft Entra Hybrid joined device selected'"
        "C004 - Require Password Change and MFA for High Risk Users" = "Scope: All Users, Target Resources: All rescources, Conditions: User Risk = High, Access Control: Grant Access when 'Require Password Change', Session: Sign-in Frequency - Every time"
        "C005 - Require MFA for Risky Sign-Ins" = "Scope: All Users, Target Resources: All resources, Conditions: Sign-in Risk = High/Medium, Access Control: Grant Access when 'Require multifactor authentication'"
    }
    
    return $descriptions[$PolicyName] -or "Standard conditional access policy configuration"
}

function Get-PolicyMappedName {
    param([string]$PolicyName)
    
    $mappings = @{
        "C001 - Block Legacy Authentication All Apps" = "Block Legacy Auth"
        "C002 - MFA Required for All Users" = "MFA for all Users"
        "C003 - Block Non Corporate Devices" = "Non Corporate Device Block"
        "C004 - Require Password Change and MFA for High Risk Users" = "User Risk"
        "C005 - Require MFA for Risky Sign-Ins" = "Risky Sign-ins"
    }
    
    return $mappings[$PolicyName] -or $PolicyName
}