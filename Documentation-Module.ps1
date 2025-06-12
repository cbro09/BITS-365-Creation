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
        # Method 1: Use ImportExcel module functions directly
        Write-LogMessage -Message "Using ImportExcel module to update spreadsheet..." -Type Info
        
        # Update M365 Due Diligence sheet - Replace company name placeholder
        try {
            $dueDilData = Import-Excel -Path $ExcelPath -WorksheetName "M365 Due Dilligence" -NoHeader
            if ($dueDilData -and $dueDilData.Count -gt 6) {
                # Find and replace company name in row 7 (index 6)
                for ($i = 0; $i -lt $dueDilData[6].PSObject.Properties.Count; $i++) {
                    $propName = "P$($i + 1)"
                    if ($dueDilData[6].$propName -and $dueDilData[6].$propName.ToString().Contains("{CompanyName}")) {
                        $dueDilData[6].$propName = $dueDilData[6].$propName.ToString().Replace("{CompanyName}", $TenantData.TenantInfo.TenantName)
                        
                        # Export back to Excel
                        $dueDilData | Export-Excel -Path $ExcelPath -WorksheetName "M365 Due Dilligence" -NoHeader -AutoSize
                        Write-LogMessage -Message "Updated company name in Due Diligence sheet" -Type Success -LogOnly
                        break
                    }
                }
            }
        }
        catch {
            Write-LogMessage -Message "Could not update Due Diligence sheet: $($_.Exception.Message)" -Type Warning -LogOnly
        }
        
        # Update Users sheet with actual user data
        if ($TenantData.Users.Count -gt 0) {
            try {
                # Create user data array for export
                $userData = @()
                $userData += [PSCustomObject]@{
                    'First Name' = 'First Name'
                    'Last Name' = 'Last Name'
                    'Email' = 'Email'
                    'Job Title' = 'Job Title'
                    'Manager email' = 'Manager email'
                    'Department' = 'Department'
                    'Office location' = 'Office location'
                    'Phone Number' = 'Phone Number'
                }
                
                foreach ($user in $TenantData.Users) {
                    $userData += [PSCustomObject]@{
                        'First Name' = $user.FirstName
                        'Last Name' = $user.LastName
                        'Email' = $user.Email
                        'Job Title' = $user.JobTitle
                        'Manager email' = $user.ManagerEmail
                        'Department' = $user.Department
                        'Office location' = $user.OfficeLocation
                        'Phone Number' = $user.PhoneNumber
                    }
                }
                
                # Export to Users sheet starting at row 6
                $userData | Export-Excel -Path $ExcelPath -WorksheetName "Users" -StartRow 6 -AutoSize -TableName "UsersTable" -TableStyle Medium2
                Write-LogMessage -Message "Updated Users sheet with $($TenantData.Users.Count) users" -Type Success -LogOnly
            }
            catch {
                Write-LogMessage -Message "Could not update Users sheet: $($_.Exception.Message)" -Type Warning -LogOnly
            }
        }
        
        # Update Licensing sheet
        if ($TenantData.Users.Count -gt 0) {
            try {
                # Create licensing data array
                $licensingData = @()
                $licensingData += [PSCustomObject]@{
                    'User Name' = 'User Name'
                    'Base License Type' = 'Base License Type'
                    'Additional Software 1' = 'Additional Software 1'
                    'Additional Software 2' = 'Additional Software 2'
                }
                
                foreach ($user in $TenantData.Users) {
                    if ($user.LicenseType) {
                        $licensingData += [PSCustomObject]@{
                            'User Name' = $user.Email
                            'Base License Type' = $user.LicenseType
                            'Additional Software 1' = ''
                            'Additional Software 2' = ''
                        }
                    }
                }
                
                if ($licensingData.Count -gt 1) {
                    # Export to Licensing sheet starting at row 7
                    $licensingData | Export-Excel -Path $ExcelPath -WorksheetName "Licensing" -StartRow 7 -StartColumn 2 -AutoSize
                    Write-LogMessage -Message "Updated Licensing sheet with user license information" -Type Success -LogOnly
                }
            }
            catch {
                Write-LogMessage -Message "Could not update Licensing sheet: $($_.Exception.Message)" -Type Warning -LogOnly
            }
        }
        
        # Add summary information using ImportExcel
        try {
            $summaryData = @()
            $summaryData += [PSCustomObject]@{
                'Summary' = 'TENANT CONFIGURATION SUMMARY'
                'Value' = ''
            }
            $summaryData += [PSCustomObject]@{
                'Summary' = "Generated: $($TenantData.TenantInfo.GeneratedDate)"
                'Value' = ''
            }
            $summaryData += [PSCustomObject]@{
                'Summary' = "Tenant Name: $($TenantData.TenantInfo.TenantName)"
                'Value' = ''
            }
            $summaryData += [PSCustomObject]@{
                'Summary' = "Default Domain: $($TenantData.TenantInfo.DefaultDomain)"
                'Value' = ''
            }
            $summaryData += [PSCustomObject]@{
                'Summary' = "Admin Email: $($TenantData.TenantInfo.AdminEmail)"
                'Value' = ''
            }
            $summaryData += [PSCustomObject]@{
                'Summary' = ''
                'Value' = ''
            }
            $summaryData += [PSCustomObject]@{
                'Summary' = 'CONFIGURATION COUNTS:'
                'Value' = ''
            }
            $summaryData += [PSCustomObject]@{
                'Summary' = "Users Created: $($TenantData.Users.Count)"
                'Value' = ''
            }
            $summaryData += [PSCustomObject]@{
                'Summary' = "Groups Created: $($TenantData.Groups.Count)"
                'Value' = ''
            }
            $summaryData += [PSCustomObject]@{
                'Summary' = "CA Policies: $($TenantData.ConditionalAccessPolicies.Count)"
                'Value' = ''
            }
            $summaryData += [PSCustomObject]@{
                'Summary' = "SharePoint Sites: $($TenantData.SharePointSites.Count)"
                'Value' = ''
            }
            
            # Add summary to Due Diligence sheet
            $summaryData | Export-Excel -Path $ExcelPath -WorksheetName "M365 Due Dilligence" -StartRow 35 -StartColumn 2 -AutoSize
            Write-LogMessage -Message "Added tenant summary information to documentation" -Type Success -LogOnly
        }
        catch {
            Write-LogMessage -Message "Could not add summary information: $($_.Exception.Message)" -Type Warning -LogOnly
        }
        
        Write-LogMessage -Message "Excel file updated and saved successfully" -Type Success
        
    } catch {
        Write-LogMessage -Message "Error updating Excel file: $($_.Exception.Message)" -Type Error
        throw
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