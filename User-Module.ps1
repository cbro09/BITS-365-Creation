# === Users.ps1 ===
# User creation and management functions

# === Excel File Validation Functions ===
function Test-ExcelFile {
    param (
        [string]$Path
    )
    
    try {
        $excel = Open-ExcelPackage -Path $Path -ErrorAction Stop
        $worksheets = $excel.Workbook.Worksheets.Name
        
        if ($worksheets -notcontains "in") {
            Close-ExcelPackage $excel -NoSave
            return @{
                Success = $false
                Message = "Excel file doesn't contain the required 'in' worksheet. Available: $($worksheets -join ', ')"
            }
        }
        
        # Import all data but filter out empty rows later
        $allData = Import-Excel -ExcelPackage $excel -WorksheetName "in"
        Close-ExcelPackage $excel -NoSave
        
        # No data at all
        if ($allData.Count -eq 0) {
            return @{
                Success = $false
                Message = "No data found in the 'in' worksheet."
            }
        }
        
        # Filter out rows that don't have all required fields
        $validData = @()
        foreach ($row in $allData) {
            if (Test-RowHasData -Row $row -RequiredColumns @('UserPrincipalName', 'DisplayName', 'Password')) {
                $validData += $row
            }
        }
        
        if ($validData.Count -eq 0) {
            return @{
                Success = $false
                Message = "No valid rows found with all required fields: UserPrincipalName, DisplayName, Password"
            }
        }
        
        # Check for required columns
        $missingColumns = @()
        $requiredColumns = @('UserPrincipalName', 'DisplayName', 'Password')
        foreach ($requiredCol in $requiredColumns) {
            if ($allData[0].PSObject.Properties.Name -notcontains $requiredCol) {
                $missingColumns += $requiredCol
            }
        }
        
        if ($missingColumns.Count -gt 0) {
            return @{
                Success = $false
                Message = "Missing required columns: $($missingColumns -join ', ')"
            }
        }
        
        return @{
            Success = $true
            Message = "Excel file is valid with $($validData.Count) user(s)."
            Data = $validData
        }
    }
    catch {
        return @{
            Success = $false
            Message = "Error analyzing Excel file: $_"
        }
    }
}

function Test-RowHasData {
    param (
        [Parameter(Mandatory = $true)]
        [object]$Row,
        
        [Parameter(Mandatory = $false)]
        [string[]]$RequiredColumns = @('UserPrincipalName', 'DisplayName', 'Password')
    )
    
    # Check if all required columns have values
    foreach ($column in $RequiredColumns) {
        if (-not (Test-NotEmpty -Value $Row.$column)) {
            return $false
        }
    }
    
    return $true
}

# === User Display Functions ===
function Show-UserList {
    param (
        [array]$Users
    )
    
    # Add null checking at the start
    if (-not $Users -or $Users.Count -eq 0) {
        Write-Host "No users found to display" -ForegroundColor Red
        return $false
    }
    
    Clear-Host
    Show-Banner
    Write-Host "== User Preview (Total: $($Users.Count)) ==" -ForegroundColor Yellow
    Write-Host ""
    
    # Create a formatted table
    $table = @()
    $table += "+-------------------------------------------------------------------+"
    $table += "| No. | UserPrincipalName                | DisplayName         | Department   |"
    $table += "+-------------------------------------------------------------------+"
    
    try {
        for ($i = 0; $i -lt [Math]::Min($Users.Count, 15); $i++) {
            # Add null checking for each user
            if (-not $Users[$i]) {
                continue
            }
            
            # Safely get values with fallbacks and null checking
            $upn = ""
            $displayName = ""
            $department = ""
            
            try {
                $upn = Get-SafeString -Value $Users[$i].UserPrincipalName -MaxLength 35 -DefaultValue "<MISSING>"
                $displayName = Get-SafeString -Value $Users[$i].DisplayName -MaxLength 20 -DefaultValue "<MISSING>"
                $department = Get-SafeString -Value $Users[$i].Department -MaxLength 13 -DefaultValue ""
            }
            catch {
                # If we can't get the values, skip this user
                continue
            }
            
            # Format strings to fixed width
            $upnStr = $upn.PadRight(35).Substring(0, 35)
            $displayNameStr = $displayName.PadRight(20).Substring(0, 20)
            $departmentStr = $department.PadRight(13).Substring(0, 13)
            
            $table += ("| {0:D3} | {1} | {2} | {3} |" -f ($i + 1), $upnStr, $displayNameStr, $departmentStr)
        }
    }
    catch {
        Write-Host "Error displaying user list: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
    
    if ($Users.Count -gt 15) {
        $table += "| ... and $($Users.Count - 15) more users ...                                  |"
    }
    
    $table += "+-------------------------------------------------------------------+"
    $table | ForEach-Object { Write-Host $_ }
    
    Write-Host ""
    $confirmation = Read-Host "Do you want to proceed with creating these users? (Y/N)"
    
    return $confirmation -eq 'Y' -or $confirmation -eq 'y'
}

# === User Creation Functions ===
function Create-M365Users {
    param (
        [array]$Users
    )
    
    $results = @{
        Success = @()
        Failed = @()
        Skipped = @()
        ManagersSet = @()
        ManagersFailed = @()
        LicensesSet = @()
        LicensesFailed = @()
    }
    
    $managerAssignments = @()
    $totalUsers = $Users.Count
    $currentUser = 0
    
    foreach ($user in $Users) {
        $currentUser++
        $statusMessage = "Processing user $currentUser of $totalUsers : $($user.DisplayName)"
        Show-Progress -Current $currentUser -Total $totalUsers -Status $statusMessage
        
        # Enhanced validation with debugging
        $upnValid = Test-NotEmpty -Value $user.UserPrincipalName
        $displayNameValid = Test-NotEmpty -Value $user.DisplayName
        $passwordValid = Test-NotEmpty -Value $user.Password
        
        # Debug output for failed user
        if (-not $upnValid -or -not $displayNameValid -or -not $passwordValid) {
            Write-LogMessage -Message "Validation failed for user: $($user.DisplayName)" -Type Warning -LogOnly
            Write-LogMessage -Message "UPN Valid: $upnValid, DisplayName Valid: $displayNameValid, Password Valid: $passwordValid" -Type Warning -LogOnly
            Write-LogMessage -Message "Password value: '$($user.Password)'" -Type Warning -LogOnly
            $results.Skipped += $user.DisplayName
            continue
        }
        
        # Additional password validation - ensure it's not just whitespace and has minimum length
        $passwordString = [string]$user.Password
        if ([string]::IsNullOrWhiteSpace($passwordString) -or $passwordString.Trim().Length -lt 1) {
            Write-LogMessage -Message "Skipping user $($user.DisplayName) - invalid password (empty or whitespace)" -Type Warning -LogOnly
            $results.Skipped += $user.DisplayName
            continue
        }
        
        # Check if user already exists
        try {
            $existingUser = Get-MgUser -Filter "UserPrincipalName eq '$($user.UserPrincipalName)'" -ErrorAction SilentlyContinue
            
            if ($existingUser) {
                Write-LogMessage -Message "User $($user.UserPrincipalName) already exists. Skipping." -Type Warning -LogOnly
                $results.Skipped += $user.DisplayName
                continue
            }
        }
        catch {
            # Continue if the user doesn't exist (which is what we want)
        }
        
        # Create password profile with explicit string conversion
        $passwordProfile = @{
            Password = [string]$user.Password.ToString().Trim()
            ForceChangePasswordNextSignIn = $true
        }
        
        # Create user parameters - only include non-empty properties
        $userParams = @{
            UserPrincipalName = [string]$user.UserPrincipalName
            DisplayName = [string]$user.DisplayName
            PasswordProfile = $passwordProfile
            MailNickName = ([string]$user.UserPrincipalName).Split("@")[0]
            AccountEnabled = $true
        }
        
        # Add optional parameters only if they exist
        if (Test-NotEmpty -Value $user.FirstName) { $userParams.GivenName = $user.FirstName }
        if (Test-NotEmpty -Value $user.LastName) { $userParams.Surname = $user.LastName }
        if (Test-NotEmpty -Value $user.JobTitle) { $userParams.JobTitle = $user.JobTitle }
        if (Test-NotEmpty -Value $user.Department) { $userParams.Department = $user.Department }
        if (Test-NotEmpty -Value $user.PhoneNumber) { $userParams.MobilePhone = $user.PhoneNumber }
        if (Test-NotEmpty -Value $user.UsageLocation) { $userParams.UsageLocation = $user.UsageLocation }
        if (Test-NotEmpty -Value $user.OfficeLocation) { $userParams.OfficeLocation = $user.OfficeLocation }
        
        # Create user
        try {
            # Debug logging before user creation
            Write-LogMessage -Message "Attempting to create user: $($user.DisplayName) with UPN: $($user.UserPrincipalName)" -Type Info -LogOnly
            Write-LogMessage -Message "Password length: $($passwordProfile.Password.Length) characters" -Type Info -LogOnly
            
            $newUser = New-MgUser @userParams
            Write-LogMessage -Message "Created user: $($user.DisplayName)" -Type Success -LogOnly
            $results.Success += $user.DisplayName
            
            # Store manager assignment for later
            if (Test-NotEmpty -Value $user.Manager) {
                $managerAssignments += @{
                    UserId = $newUser.Id
                    UserPrincipalName = $user.UserPrincipalName
                    DisplayName = $user.DisplayName
                    ManagerUPN = $user.Manager
                }
            }
            
            # Set license extension attribute
            if (Test-NotEmpty -Value $user.License) {
                try {
                    $extensionAttributes = @{
                        "onPremisesExtensionAttributes" = @{
                            "extensionAttribute1" = $user.License
                        }
                    }
                    
                    Update-MgUser -UserId $newUser.Id -AdditionalProperties $extensionAttributes
                    Write-LogMessage -Message "Set license attribute for $($user.DisplayName)" -Type Success -LogOnly
                    $results.LicensesSet += $user.DisplayName
                }
                catch {
                    Write-LogMessage -Message "Failed to set license attribute for $($user.DisplayName) - $($_.Exception.Message)" -Type Error -LogOnly
                    $results.LicensesFailed += $user.DisplayName
                }
            }
        }
        catch {
            Write-LogMessage -Message "Error creating $($user.DisplayName) - $($_.Exception.Message)" -Type Error -LogOnly
            $results.Failed += $user.DisplayName
        }
    }
    
    Write-Host ""
    Write-LogMessage -Message "Setting manager relationships..." -Type Info
    
    $currentManager = 0
    $totalManagers = $managerAssignments.Count
    
    foreach ($assignment in $managerAssignments) {
        $currentManager++
        Show-Progress -Current $currentManager -Total $totalManagers -Status "Setting manager for $($assignment.DisplayName)"
        
        try {
            $manager = Get-MgUser -Filter "UserPrincipalName eq '$($assignment.ManagerUPN)'" -ErrorAction Stop
            
            if ($manager) {
                Set-MgUserManagerByRef -UserId $assignment.UserId -BodyParameter @{
                    "@odata.id" = "https://graph.microsoft.com/v1.0/users/$($manager.Id)"
                }
                Write-LogMessage -Message "Set manager for $($assignment.DisplayName)" -Type Success -LogOnly
                $results.ManagersSet += $assignment.DisplayName
            }
            else {
                Write-LogMessage -Message "Manager with UPN $($assignment.ManagerUPN) not found for user $($assignment.DisplayName)" -Type Warning -LogOnly
                $results.ManagersFailed += $assignment.DisplayName
            }
        }
        catch {
            Write-LogMessage -Message "Failed to set manager for $($assignment.DisplayName) - $($_.Exception.Message)" -Type Warning -LogOnly
            $results.ManagersFailed += $assignment.DisplayName
        }
    }
    
    Write-Host ""
    return $results
}

# === Results Display Functions ===
function Show-Results {
    param (
        [hashtable]$Results
    )
    
    Clear-Host
    Show-Banner
    Write-Host "== Operation Results ==" -ForegroundColor Yellow
    Write-Host ""
    
    $table = @()
    $table += "+-------------------------------------+"
    $table += "| Operation                  | Count  |"
    $table += "+-------------------------------------+"
    $table += ("| Users Created              | {0,-6} |" -f $Results.Success.Count)
    $table += ("| Users Failed               | {0,-6} |" -f $Results.Failed.Count)
    $table += ("| Users Skipped              | {0,-6} |" -f $Results.Skipped.Count)
    $table += ("| Managers Set               | {0,-6} |" -f $Results.ManagersSet.Count)
    $table += ("| Managers Failed            | {0,-6} |" -f $Results.ManagersFailed.Count)
    $table += ("| License Attributes Set     | {0,-6} |" -f $Results.LicensesSet.Count)
    $table += ("| License Attributes Failed  | {0,-6} |" -f $Results.LicensesFailed.Count)
    $table += "+-------------------------------------+"
    $table | ForEach-Object { Write-Host $_ }
    
    Write-Host ""
    Write-LogMessage -Message "Log file saved to: $($config.LogFile)" -Type Info
    
    if ($Results.Failed.Count -gt 0 -or $Results.ManagersFailed.Count -gt 0 -or $Results.LicensesFailed.Count -gt 0) {
        Write-Host "Some operations failed. See log file for details." -ForegroundColor Yellow
    }
    
    Write-Host ""
    return Read-Host "Would you like to export detailed results to Excel? (Y/N)"
}

function Export-ResultsToExcel {
    param (
        [hashtable]$Results,
        [string]$ExcelPath
    )
    
    try {
        $exportPath = [System.IO.Path]::Combine([System.IO.Path]::GetDirectoryName($ExcelPath), "UserCreationResults_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx")
        
        # Create objects for each result category
        $successUsers = $Results.Success | ForEach-Object { 
            [PSCustomObject]@{
                DisplayName = $_
                Status = "Success" 
            }
        }
        
        $failedUsers = $Results.Failed | ForEach-Object { 
            [PSCustomObject]@{
                DisplayName = $_
                Status = "Failed" 
            }
        }
        
        $skippedUsers = $Results.Skipped | ForEach-Object { 
            [PSCustomObject]@{
                DisplayName = $_
                Status = "Skipped"
            }
        }
        
        $allUsers = @($successUsers) + @($failedUsers) + @($skippedUsers)
        
        $managerResults = @()
        foreach ($user in $Results.ManagersSet) {
            $managerResults += [PSCustomObject]@{
                DisplayName = $user
                Operation = "Set Manager"
                Status = "Success"
            }
        }
        
        foreach ($user in $Results.ManagersFailed) {
            $managerResults += [PSCustomObject]@{
                DisplayName = $user
                Operation = "Set Manager"
                Status = "Failed"
            }
        }
        
        $licenseResults = @()
        foreach ($user in $Results.LicensesSet) {
            $licenseResults += [PSCustomObject]@{
                DisplayName = $user
                Operation = "Set License Attribute"
                Status = "Success"
            }
        }
        
        foreach ($user in $Results.LicensesFailed) {
            $licenseResults += [PSCustomObject]@{
                DisplayName = $user
                Operation = "Set License Attribute"
                Status = "Failed"
            }
        }
        
        # Export to Excel
        $allUsers | Export-Excel -Path $exportPath -WorksheetName "User Creation" -AutoSize -TableName "UserCreation" -TableStyle Medium2
        
        if ($managerResults.Count -gt 0) {
            $managerResults | Export-Excel -Path $exportPath -WorksheetName "Manager Assignments" -AutoSize -TableName "ManagerAssignments" -TableStyle Medium2
        }
        
        if ($licenseResults.Count -gt 0) {
            $licenseResults | Export-Excel -Path $exportPath -WorksheetName "License Attributes" -AutoSize -TableName "LicenseAttributes" -TableStyle Medium2
        }
        
        # Add summary worksheet
        $summary = @(
            [PSCustomObject]@{ Operation = "Users Created"; Count = $Results.Success.Count },
            [PSCustomObject]@{ Operation = "Users Failed"; Count = $Results.Failed.Count },
            [PSCustomObject]@{ Operation = "Users Skipped"; Count = $Results.Skipped.Count },
            [PSCustomObject]@{ Operation = "Managers Set"; Count = $Results.ManagersSet.Count },
            [PSCustomObject]@{ Operation = "Managers Failed"; Count = $Results.ManagersFailed.Count },
            [PSCustomObject]@{ Operation = "License Attributes Set"; Count = $Results.LicensesSet.Count },
            [PSCustomObject]@{ Operation = "License Attributes Failed"; Count = $Results.LicensesFailed.Count }
        )
        
        $summary | Export-Excel -Path $exportPath -WorksheetName "Summary" -AutoSize -TableName "Summary" -TableStyle Medium2
        
        Write-LogMessage -Message "Results exported to: $exportPath" -Type Success
        return $true
    }
    catch {
        Write-LogMessage -Message "Failed to export results: $_" -Type Error
        return $false
    }
}

# === Debug Functions ===
function Debug-ExcelData {
    Write-Host "=== DEBUGGING EXCEL DATA ===" -ForegroundColor Yellow
    
    # Use file explorer to select Excel file
    Add-Type -AssemblyName System.Windows.Forms
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Title = "Select Excel File to Debug"
    $openFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
    $openFileDialog.InitialDirectory = "$env:USERPROFILE\Documents"
    
    if ($openFileDialog.ShowDialog() -eq 'OK') {
        $ExcelPath = $openFileDialog.FileName
        Write-Host "Selected file: $ExcelPath" -ForegroundColor Cyan
    } else {
        Write-Host "File selection canceled" -ForegroundColor Yellow
        return
    }
    
    try {
        # Import raw data
        $rawData = Import-Excel -Path $ExcelPath -WorksheetName "in"
        
        Write-Host "Total rows imported: $($rawData.Count)" -ForegroundColor Cyan
        Write-Host "Column names: $($rawData[0].PSObject.Properties.Name -join ', ')" -ForegroundColor Cyan
        
        # Check each user's data
        for ($i = 0; $i -lt $rawData.Count; $i++) {
            $user = $rawData[$i]
            Write-Host "`n--- User $($i + 1): $($user.DisplayName) ---" -ForegroundColor Green
            Write-Host "UserPrincipalName: '$($user.UserPrincipalName)' (Type: $($user.UserPrincipalName.GetType().Name))" -ForegroundColor White
            Write-Host "DisplayName: '$($user.DisplayName)' (Type: $($user.DisplayName.GetType().Name))" -ForegroundColor White
            Write-Host "Password: '$($user.Password)' (Type: $($user.Password.GetType().Name))" -ForegroundColor White
            Write-Host "Password Length: $($user.Password.ToString().Length)" -ForegroundColor White
            Write-Host "Password IsNull: $($user.Password -eq $null)" -ForegroundColor White
            Write-Host "Password IsEmpty: $([string]::IsNullOrEmpty($user.Password))" -ForegroundColor White
            Write-Host "Password IsWhitespace: $([string]::IsNullOrWhiteSpace([string]$user.Password))" -ForegroundColor White
        }
        
        Write-Host "`n=== END DEBUG ===" -ForegroundColor Yellow
    }
    catch {
        Write-Host "Error reading Excel file: $_" -ForegroundColor Red
    }
}

# === Main User Creation Function ===
function New-TenantUsers {
    Write-LogMessage -Message "Starting user creation process..." -Type Info
    
    # COMPLETE module reset to match working script exactly
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
        
        # Force load ONLY the exact modules from working script in exact order
        $userCreationModules = @('Microsoft.Graph.Users', 'Microsoft.Graph.Identity.DirectoryManagement', 'ImportExcel')
        
        Write-LogMessage -Message "Loading ONLY user creation modules in exact order..." -Type Info
        foreach ($module in $userCreationModules) {
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
        
        # Connect with EXACT scopes from working script
        $userCreationScopes = @("User.ReadWrite.All", "Directory.ReadWrite.All")
        
        Write-LogMessage -Message "Connecting to Microsoft Graph with user creation scopes only..." -Type Info
        Connect-MgGraph -Scopes $userCreationScopes -ErrorAction Stop | Out-Null
        
        $context = Get-MgContext
        Write-LogMessage -Message "Connected to Microsoft Graph as $($context.Account)" -Type Success
        Write-LogMessage -Message "Active scopes: $($context.Scopes -join ', ')" -Type Info -LogOnly
        
        # Use EXACT file handling from working script
        $defaultExcelPath = "$env:USERPROFILE\Documents\users.xlsx"
        
        $excelFile = $null
        if (Test-Path -Path $defaultExcelPath) {
            try {
                $result = Test-ExcelFile -Path $defaultExcelPath
                if ($result -and $result.Success -and $result.Data) {
                    Write-LogMessage -Message "Excel file found at default location and is valid" -Type Success
                    $excelFile = @{
                        Success = $true
                        Path = $defaultExcelPath
                        Data = $result.Data
                    }
                }
                else {
                    Write-LogMessage -Message "Excel file found but is invalid or has no data" -Type Warning
                }
            }
            catch {
                Write-LogMessage -Message "Error reading default Excel file: $($_.Exception.Message)" -Type Warning
            }
        }
        
        if (-not $excelFile -or -not $excelFile.Success) {
            # File dialog exactly like working script
            Add-Type -AssemblyName System.Windows.Forms
            $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
            $openFileDialog.Title = "Select Users Excel File"
            $openFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
            
            try {
                $openFileDialog.InitialDirectory = [System.IO.Path]::GetDirectoryName($defaultExcelPath)
            }
            catch {
                $openFileDialog.InitialDirectory = "$env:USERPROFILE\Documents"
            }
            
            if ($openFileDialog.ShowDialog() -eq 'OK') {
                try {
                    $result = Test-ExcelFile -Path $openFileDialog.FileName
                    if ($result -and $result.Success -and $result.Data) {
                        Write-LogMessage -Message "Selected Excel file is valid" -Type Success
                        $excelFile = @{
                            Success = $true
                            Path = $openFileDialog.FileName
                            Data = $result.Data
                        }
                    }
                    else {
                        Write-LogMessage -Message "Selected Excel file is invalid or has no data" -Type Error
                        return $false
                    }
                }
                catch {
                    Write-LogMessage -Message "Error reading selected Excel file: $($_.Exception.Message)" -Type Error
                    return $false
                }
            }
            else {
                Write-LogMessage -Message "File selection canceled by user" -Type Warning
                return $false
            }
        }
        
        # Validate data
        if (-not $excelFile -or -not $excelFile.Data -or $excelFile.Data.Count -eq 0) {
            Write-LogMessage -Message "No valid user data found in Excel file" -Type Error
            return $false
        }
        
        # Show user list exactly like working script
        try {
            $proceedWithCreation = Show-UserList -Users $excelFile.Data
            if (-not $proceedWithCreation) {
                Write-LogMessage -Message "User creation canceled by user" -Type Info
                return $false
            }
        }
        catch {
            Write-LogMessage -Message "Error displaying user list: $($_.Exception.Message)" -Type Error
            return $false
        }
        
        # Create users exactly like working script
        try {
            Write-LogMessage -Message "Starting user creation process..." -Type Info
            $results = Create-M365Users -Users $excelFile.Data
            
            if (-not $results) {
                Write-LogMessage -Message "User creation returned no results" -Type Error
                return $false
            }
        }
        catch {
            Write-LogMessage -Message "Error during user creation: $($_.Exception.Message)" -Type Error
            return $false
        }
        
        # Show results exactly like working script
        try {
            $exportResults = Show-Results -Results $results
            if ($exportResults -eq 'Y' -or $exportResults -eq 'y') {
                Export-ResultsToExcel -Results $results -ExcelPath $excelFile.Path
            }
        }
        catch {
            Write-LogMessage -Message "Error displaying results: $($_.Exception.Message)" -Type Error
        }
        
        Write-LogMessage -Message "User creation workflow completed" -Type Success
        return $true
    }
    catch {
        Write-LogMessage -Message "Error in user creation process - $($_.Exception.Message)" -Type Error
        return $false
    }
}