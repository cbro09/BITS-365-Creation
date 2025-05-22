#requires -Version 5.1
<#
.SYNOPSIS
    Microsoft 365 User Creation Utility
.DESCRIPTION
    Creates Microsoft 365 users from an Excel spreadsheet with an enhanced user experience.
.NOTES
    Author: Improved by Claude
    Version: 3.0
    Requirements: PowerShell 5.1, Microsoft Graph PowerShell SDK, ImportExcel module
.EXAMPLE
    .\Create-M365Users.ps1
#>

# ===== Configuration =====
$config = @{
    RequiredModules = @('Microsoft.Graph.Users', 'Microsoft.Graph.Identity.DirectoryManagement', 'ImportExcel')
    GraphScopes = @("User.ReadWrite.All", "Directory.ReadWrite.All")
    DefaultExcelPath = "$env:USERPROFILE\Documents\users.xlsx"
    LogFile = "$env:TEMP\M365UserCreation_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    RequiredColumns = @('UserPrincipalName', 'DisplayName', 'Password')
}

# ===== Helper Functions =====
function Write-LogMessage {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet('Info', 'Success', 'Warning', 'Error')]
        [string]$Type = 'Info',
        
        [Parameter(Mandatory = $false)]
        [switch]$LogOnly
    )
    
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logMessage = "[$timestamp] [$Type] $Message"
    Add-Content -Path $config.LogFile -Value $logMessage
    
    if (-not $LogOnly) {
        switch ($Type) {
            'Info'    { Write-Host "[INFO] $Message" -ForegroundColor Cyan }
            'Success' { Write-Host "[SUCCESS] $Message" -ForegroundColor Green }
            'Warning' { Write-Host "[WARNING] $Message" -ForegroundColor Yellow }
            'Error'   { Write-Host "[ERROR] $Message" -ForegroundColor Red }
        }
    }
}

function Show-Banner {
    Write-Host ""
    Write-Host "+------------------------------------------------+" -ForegroundColor Blue
    Write-Host "|     Microsoft 365 User Creation Utility        |" -ForegroundColor Magenta
    Write-Host "+------------------------------------------------+" -ForegroundColor Blue
    Write-Host ""
}

function Show-Menu {
    param (
        [string]$Title = 'Menu',
        [array]$Options
    )
    
    Clear-Host
    Show-Banner
    Write-Host "== $Title ==" -ForegroundColor Yellow
    Write-Host ""
    
    for ($i = 0; $i -lt $Options.Count; $i++) {
        Write-Host " [$($i + 1)] " -ForegroundColor Yellow -NoNewline
        Write-Host $Options[$i] -ForegroundColor White
    }
    
    Write-Host ""
    $selection = Read-Host "Enter your choice (1-$($Options.Count))"
    
    if ($selection -match '^\d+$' -and [int]$selection -ge 1 -and [int]$selection -le $Options.Count) {
        return [int]$selection
    }
    else {
        Write-Host "Invalid selection. Please try again." -ForegroundColor Red
        Start-Sleep -Seconds 1
        return Show-Menu -Title $Title -Options $Options
    }
}

function Show-Progress {
    param (
        [int]$Current,
        [int]$Total,
        [string]$Status
    )
    
    $percentComplete = [math]::Round(($Current / $Total) * 100)
    $progressBarWidth = 50
    $filledWidth = [math]::Round(($percentComplete / 100) * $progressBarWidth)
    $emptyWidth = $progressBarWidth - $filledWidth
    
    $progressBar = '[' + ('#' * $filledWidth) + (' ' * $emptyWidth) + ']'
    
    Write-Host "`r$progressBar $percentComplete% - $Status                                     " -NoNewline
}

function Test-NotEmpty {
    param (
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [AllowEmptyString()]
        [object]$Value
    )
    
    # Check if value is null or empty string
    if ($null -eq $Value -or 
        ($Value -is [string] -and [string]::IsNullOrWhiteSpace($Value)) -or
        ($Value -is [array] -and $Value.Count -eq 0) -or
        ($Value -is [System.Collections.ICollection] -and $Value.Count -eq 0)) {
        return $false
    }
    
    return $true
}

function Test-RowHasData {
    param (
        [Parameter(Mandatory = $true)]
        [object]$Row,
        
        [Parameter(Mandatory = $false)]
        [string[]]$RequiredColumns = @()
    )
    
    # If no required columns are specified, check any property
    if ($RequiredColumns.Count -eq 0) {
        foreach ($prop in $Row.PSObject.Properties) {
            if (Test-NotEmpty -Value $prop.Value) {
                return $true
            }
        }
        return $false
    }
    
    # Check if all required columns have values
    foreach ($column in $RequiredColumns) {
        if (-not (Test-NotEmpty -Value $Row.$column)) {
            return $false
        }
    }
    
    return $true
}

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
            if (Test-RowHasData -Row $row -RequiredColumns $config.RequiredColumns) {
                $validData += $row
            }
        }
        
        if ($validData.Count -eq 0) {
            return @{
                Success = $false
                Message = "No valid rows found with all required fields: $($config.RequiredColumns -join ', ')"
            }
        }
        
        # Check for required columns
        $missingColumns = @()
        foreach ($requiredCol in $config.RequiredColumns) {
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

function Get-ExcelFile {
    param (
        [string]$DefaultPath
    )
    
    if (Test-Path -Path $DefaultPath) {
        $result = Test-ExcelFile -Path $DefaultPath
        if ($result.Success) {
            Write-LogMessage -Message "Excel file found at default location and is valid" -Type Success
            return @{
                Success = $true
                Path = $DefaultPath
                Data = $result.Data
            }
        }
        else {
            Write-LogMessage -Message "Excel file found but is invalid: $($result.Message)" -Type Warning
        }
    }
    
    # Ask user for file path
    Add-Type -AssemblyName System.Windows.Forms
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Title = "Select Users Excel File"
    $openFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
    $openFileDialog.InitialDirectory = [System.IO.Path]::GetDirectoryName($DefaultPath)
    
    if ($openFileDialog.ShowDialog() -eq 'OK') {
        $result = Test-ExcelFile -Path $openFileDialog.FileName
        if ($result.Success) {
            Write-LogMessage -Message "Selected Excel file is valid" -Type Success
            return @{
                Success = $true
                Path = $openFileDialog.FileName
                Data = $result.Data
            }
        }
        else {
            Write-LogMessage -Message "Selected Excel file is invalid: $($result.Message)" -Type Error
            return @{
                Success = $false
                Message = $result.Message
            }
        }
    }
    else {
        return @{
            Success = $false
            Message = "File selection canceled by user."
        }
    }
}

function Get-SafeString {
    param (
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [AllowEmptyString()]
        [object]$Value,
        
        [Parameter(Mandatory = $false)]
        [int]$MaxLength = -1,
        
        [Parameter(Mandatory = $false)]
        [string]$DefaultValue = ""
    )
    
    # Handle null or empty
    if (-not (Test-NotEmpty -Value $Value)) {
        return $DefaultValue
    }
    
    # Convert to string
    $result = "$Value"
    
    # Truncate if needed
    if ($MaxLength -gt 0 -and $result.Length -gt $MaxLength) {
        $result = $result.Substring(0, $MaxLength)
    }
    
    return $result
}

function Show-UserList {
    param (
        [array]$Users
    )
    
    Clear-Host
    Show-Banner
    Write-Host "== User Preview (Total: $($Users.Count)) ==" -ForegroundColor Yellow
    Write-Host ""
    
    # Create a formatted table
    $table = @()
    $table += "+-------------------------------------------------------------------+"
    $table += "| No. | UserPrincipalName                | DisplayName         | Department   |"
    $table += "+-------------------------------------------------------------------+"
    
    for ($i = 0; $i -lt [Math]::Min($Users.Count, 15); $i++) {
        # Safely get values with fallbacks
        $upn = Get-SafeString -Value $Users[$i].UserPrincipalName -MaxLength 35 -DefaultValue "<MISSING>"
        $displayName = Get-SafeString -Value $Users[$i].DisplayName -MaxLength 20 -DefaultValue "<MISSING>"
        $department = Get-SafeString -Value $Users[$i].Department -MaxLength 13 -DefaultValue ""
        
        # Format strings to fixed width
        $upnStr = $upn.PadRight(35).Substring(0, 35)
        $displayNameStr = $displayName.PadRight(20).Substring(0, 20)
        $departmentStr = $department.PadRight(13).Substring(0, 13)
        
        $table += ("| {0:D3} | {1} | {2} | {3} |" -f ($i + 1), $upnStr, $displayNameStr, $departmentStr)
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

function Install-RequiredModules {
    $installedCount = 0
    $totalModules = $config.RequiredModules.Count
    
    foreach ($module in $config.RequiredModules) {
        $installedCount++
        Show-Progress -Current $installedCount -Total $totalModules -Status "Checking module: $module"
        
        try {
            if (-not (Get-Module -ListAvailable -Name $module)) {
                Write-LogMessage -Message "Installing $module module..." -Type Info
                Install-Module -Name $module -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
                Write-LogMessage -Message "$module module installed successfully" -Type Success
            }
            else {
                Write-LogMessage -Message "$module module already installed" -Type Info -LogOnly
            }
            
            # Import the module
            Import-Module -Name $module -Force -ErrorAction Stop
            Write-LogMessage -Message "$module module loaded successfully" -Type Success -LogOnly
        }
        catch {
            Write-LogMessage -Message "Failed to install/import $module module. Error: $_" -Type Error
            return $false
        }
    }
    
    Write-Host ""
    return $true
}

function Connect-ToGraph {
    try {
        # Check if we're already connected
        $graphConnection = Get-MgContext -ErrorAction SilentlyContinue
        
        if ($graphConnection) {
            Write-LogMessage -Message "Already connected to Microsoft Graph as $($graphConnection.Account)" -Type Info
            
            $reconnect = Read-Host "Do you want to reconnect with a different account? (Y/N)"
            if ($reconnect -eq 'Y' -or $reconnect -eq 'y') {
                Disconnect-MgGraph | Out-Null
            }
            else {
                return $true
            }
        }
        
        Write-LogMessage -Message "Connecting to Microsoft Graph..." -Type Info
        Connect-MgGraph -Scopes $config.GraphScopes -ErrorAction Stop | Out-Null
        
        $context = Get-MgContext
        Write-LogMessage -Message "Successfully connected to Microsoft Graph as $($context.Account)" -Type Success
        
        return $true
    }
    catch {
        Write-LogMessage -Message "Failed to connect to Microsoft Graph. Error: $_" -Type Error
        return $false
    }
}

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
        $statusMessage = "Processing user $currentUser of $totalUsers`: $($user.DisplayName)"
        Show-Progress -Current $currentUser -Total $totalUsers -Status $statusMessage
        
        # Skip if required fields are missing
        if (-not (Test-NotEmpty -Value $user.UserPrincipalName) -or 
            -not (Test-NotEmpty -Value $user.DisplayName) -or
            -not (Test-NotEmpty -Value $user.Password)) {
            
            Write-LogMessage -Message "Skipping user with missing required fields: $($user.DisplayName)" -Type Warning -LogOnly
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
        
        # Create password profile
        $passwordProfile = @{
            Password = $user.Password
            ForceChangePasswordNextSignIn = $true
        }
        
        # Create user parameters - only include non-empty properties
        $userParams = @{
            UserPrincipalName = $user.UserPrincipalName
            DisplayName = $user.DisplayName
            PasswordProfile = $passwordProfile
            MailNickName = ($user.UserPrincipalName.Split("@"))[0]
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
                    # Use AdditionalProperties instead of OnPremisesExtensionAttributes
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
                    Write-LogMessage -Message "Failed to set license attribute for $($user.DisplayName): $_" -Type Error -LogOnly
                    $results.LicensesFailed += $user.DisplayName
                }
            }
        }
        catch {
            Write-LogMessage -Message "Error creating $($user.DisplayName): $_" -Type Error -LogOnly
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
            Write-LogMessage -Message "Failed to set manager for $($assignment.DisplayName): $_" -Type Warning -LogOnly
            $results.ManagersFailed += $assignment.DisplayName
        }
    }
    
    Write-Host ""
    return $results
}

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

# ===== Main Execution =====
Clear-Host
Show-Banner

# Start logging
try {
    $logDir = [System.IO.Path]::GetDirectoryName($config.LogFile)
    if (-not (Test-Path -Path $logDir)) {
        New-Item -Path $logDir -ItemType Directory -Force | Out-Null
    }
    
    New-Item -Path $config.LogFile -ItemType File -Force | Out-Null
    Write-LogMessage -Message "Microsoft 365 User Creation Utility started" -Type Info
    Write-LogMessage -Message "PowerShell version: $($PSVersionTable.PSVersion)" -Type Info -LogOnly
    Write-LogMessage -Message "Computer name: $env:COMPUTERNAME" -Type Info -LogOnly
    Write-LogMessage -Message "User context: $env:USERNAME" -Type Info -LogOnly
}
catch {
    Write-Host "Failed to start logging: $_" -ForegroundColor Red
    exit
}

# Main menu loop
$exitScript = $false
while (-not $exitScript) {
    $choice = Show-Menu -Title "Main Menu" -Options @(
        "Create Users from Excel"
        "View Log File"
        "Exit"
    )
    
    switch ($choice) {
        1 {
            # Create users workflow
            Write-LogMessage -Message "Starting user creation workflow" -Type Info
            
            # Install required modules
            Write-LogMessage -Message "Checking required modules..." -Type Info
            $modulesInstalled = Install-RequiredModules
            if (-not $modulesInstalled) {
                Write-LogMessage -Message "Failed to install required modules. Please fix the issues and try again." -Type Error
                Read-Host "Press Enter to continue"
                continue
            }
            
            # Connect to Microsoft Graph
            $graphConnected = Connect-ToGraph
            if (-not $graphConnected) {
                Write-LogMessage -Message "Failed to connect to Microsoft Graph. Please try again." -Type Error
                Read-Host "Press Enter to continue"
                continue
            }
            
            # Get Excel file
            $excelFile = Get-ExcelFile -DefaultPath $config.DefaultExcelPath
            if (-not $excelFile.Success) {
                Write-LogMessage -Message "Excel file process failed: $($excelFile.Message)" -Type Error
                Read-Host "Press Enter to continue"
                continue
            }
            
            # Show user list and confirm
            $proceedWithCreation = Show-UserList -Users $excelFile.Data
            if (-not $proceedWithCreation) {
                Write-LogMessage -Message "User creation canceled by user" -Type Info
                Read-Host "Press Enter to continue"
                continue
            }
            
            # Create users
            Write-LogMessage -Message "Starting user creation process..." -Type Info
            $results = Create-M365Users -Users $excelFile.Data
            
            # Show results and export if requested
            $exportResults = Show-Results -Results $results
            if ($exportResults -eq 'Y' -or $exportResults -eq 'y') {
                Export-ResultsToExcel -Results $results -ExcelPath $excelFile.Path
            }
            
            # Disconnect from Graph
            Disconnect-MgGraph | Out-Null
            Write-LogMessage -Message "Disconnected from Microsoft Graph" -Type Info
            
            Read-Host "Press Enter to continue"
        }
        2 {
            # View log file
            if (Test-Path -Path $config.LogFile) {
                try {
                    # Open log file with default text editor
                    Start-Process -FilePath $config.LogFile
                }
                catch {
                    Write-LogMessage -Message "Failed to open log file: $_" -Type Error
                }
            }
            else {
                Write-LogMessage -Message "Log file not found" -Type Warning
            }
            
            Read-Host "Press Enter to continue"
        }
        3 {
            # Exit
            $exitScript = $true
            Write-LogMessage -Message "Microsoft 365 User Creation Utility ended" -Type Info
        }
    }
}

# Final cleanup
if (Get-MgContext) {
    Disconnect-MgGraph | Out-Null
}

Write-Host ""
Write-Host "Thank you for using the Microsoft 365 User Creation Utility!" -ForegroundColor Cyan
Start-Sleep -Seconds 1