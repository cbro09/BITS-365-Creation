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
}#requires -Version 5.1
<#
.SYNOPSIS
    Unified Microsoft 365 Tenant Setup Utility
.DESCRIPTION
    Complete Microsoft 365 tenant setup automation including groups, CA policies, SharePoint, Intune and users
.NOTES
    Version: 1.0
    Requirements: PowerShell 5.1 or later, Microsoft Graph PowerShell SDK
#>

# === Configuration ===
$config = @{
    # Core configuration
    LogFile = "$env:TEMP\M365TenantSetup_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    RequiredModules = @(
        'Microsoft.Graph.Authentication',
        'Microsoft.Graph.Identity.DirectoryManagement',
        'Microsoft.Graph.Users',
        'Microsoft.Graph.Groups',
        'Microsoft.Graph.DeviceManagement',
        'Microsoft.Online.SharePoint.PowerShell',
        'ImportExcel'
    )
    GraphScopes = @(
        "User.ReadWrite.All", 
        "Group.ReadWrite.All", 
        "Directory.ReadWrite.All", 
        "Policy.ReadWrite.ConditionalAccess",
        "DeviceManagementConfiguration.ReadWrite.All",
        "DeviceManagementManagedDevices.ReadWrite.All",
        "DeviceManagementApps.ReadWrite.All"
    )
    # Default resource names and settings
    DefaultGroups = @{
        Security = @("BITS Admin", "SSPR Enabled", "NoMFA Exemption")
        License = @("BusinessBasic", "BusinessStandard", "BusinessPremium", "ExchangeOnline1", "ExchangeOnline2")
    }
    SharePoint = @{
        SiteTemplate = "SITEPAGEPUBLISHING#0"
        DefaultSites = @("HR", "Processes", "Templates", "Documents", "Policies")
        StorageQuota = 1024
    }
}

# === Helper Functions ===
function Test-NotEmpty {
    param (
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [AllowEmptyString()]
        [object]$Value
    )
    
    if ($null -eq $Value -or 
        ($Value -is [string] -and [string]::IsNullOrWhiteSpace($Value)) -or
        ($Value -is [array] -and $Value.Count -eq 0) -or
        ($Value -is [System.Collections.ICollection] -and $Value.Count -eq 0)) {
        return $false
    }
    
    return $true
}

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
    Write-Host "|      Unified Microsoft 365 Tenant Setup        |" -ForegroundColor Magenta
    Write-Host "+------------------------------------------------+" -ForegroundColor Blue
    Write-Host ""
    Write-Host "IMPORTANT: Ensure you have Global Administrator" -ForegroundColor Red
    Write-Host "credentials for the target Microsoft 365 tenant" -ForegroundColor Red
    Write-Host "before proceeding with this script." -ForegroundColor Red
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
    Write-Host "`r$progressBar $percentComplete% - $Status                              " -NoNewline
}

function Install-RequiredModules {
    Write-LogMessage -Message "Checking and installing required modules..." -Type Info
    $moduleCount = $config.RequiredModules.Count
    $currentModule = 0
    
    foreach ($moduleName in $config.RequiredModules) {
        $currentModule++
        Show-Progress -Current $currentModule -Total $moduleCount -Status "Processing module: $moduleName"
        
        try {
            if (-not (Get-Module -ListAvailable -Name $moduleName)) {
                Write-LogMessage -Message "Installing $moduleName module..." -Type Info -LogOnly
                Install-Module -Name $moduleName -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
                Write-LogMessage -Message "$moduleName module installed successfully" -Type Success -LogOnly
            }
            else {
                Write-LogMessage -Message "$moduleName module already installed" -Type Info -LogOnly
            }
            
            # Import the module
            Import-Module -Name $moduleName -Force -ErrorAction Stop
            Write-LogMessage -Message "$moduleName module imported successfully" -Type Success -LogOnly
        }
        catch {
            Write-LogMessage -Message "Failed to install/import $moduleName module - $($_.Exception.Message)" -Type Error
            return $false
        }
    }
    
    Write-Host ""
    return $true
}

function Import-RequiredGraphModules {
    # Graph modules are now installed individually, so just import them
    $graphModules = @(
        'Microsoft.Graph.Authentication',
        'Microsoft.Graph.Identity.DirectoryManagement',
        'Microsoft.Graph.Users',
        'Microsoft.Graph.Groups',
        'Microsoft.Graph.DeviceManagement'
    )
    
    foreach ($module in $graphModules) {
        try {
            Import-Module $module -ErrorAction Stop
        }
        catch {
            Write-LogMessage -Message "Failed to import $module - $($_.Exception.Message)" -Type Warning
        }
    }
}

function Connect-ToGraphAndVerify {
    try {
        Import-RequiredGraphModules
        
        # Check if already connected
        $graphConnection = Get-MgContext -ErrorAction SilentlyContinue
        
        if ($graphConnection) {
            Write-LogMessage -Message "Already connected to Microsoft Graph as $($graphConnection.Account)" -Type Info
            
            $reconnect = Read-Host "Do you want to reconnect with a different account? (Y/N)"
            if ($reconnect -eq 'Y' -or $reconnect -eq 'y') {
                Disconnect-MgGraph | Out-Null
                # Proceed to connect with new account
            }
            else {
                # Still connected but we MUST verify domain for multiple tenants
                Write-LogMessage -Message "Verifying current tenant domain..." -Type Info
                $verified = Test-TenantDomain
                if (-not $verified) {
                    Write-LogMessage -Message "Domain verification failed. Please disconnect and connect to the correct tenant." -Type Error
                    return $false
                }
                return $true
            }
        }
        
        Write-LogMessage -Message "Connecting to Microsoft Graph..." -Type Info
        Write-LogMessage -Message "Required scopes: $($config.GraphScopes -join ', ')" -Type Info
        
        # Connect with required scopes
        Connect-MgGraph -Scopes $config.GraphScopes -ErrorAction Stop | Out-Null
        
        $context = Get-MgContext
        Write-LogMessage -Message "Successfully connected to Microsoft Graph as $($context.Account)" -Type Success
        
        # ALWAYS verify tenant domain - critical for multiple tenant scenarios
        Write-LogMessage -Message "Verifying tenant domain..." -Type Info
        $verified = Test-TenantDomain
        if (-not $verified) {
            Write-LogMessage -Message "Domain verification failed. Please connect to the correct tenant." -Type Error
            Disconnect-MgGraph | Out-Null
            return $false
        }
        
        return $true
    }
    catch {
        Write-LogMessage -Message "Failed to connect to Microsoft Graph - $($_.Exception.Message)" -Type Error
        return $false
    }
}

function Test-TenantDomain {
    try {
        # Get organization details
        $organization = Get-MgOrganization
        $verifiedDomains = $organization.VerifiedDomains
        $defaultDomain = $verifiedDomains | Where-Object { $_.IsDefault -eq $true }
        
        Write-Host "Current default domain: $($defaultDomain.Name)" -ForegroundColor Cyan
        
        $confirmation = Read-Host "Is this the correct default domain for this tenant? (Y/N)"
        if ($confirmation -ne 'Y' -and $confirmation -ne 'y') {
            return $false
        }
        
        # Save tenant state information
        $script:TenantState = @{
            DefaultDomain = $defaultDomain.Name
            TenantName = $organization.DisplayName
            TenantId = $organization.Id
            CreatedGroups = @{}
            AdminEmail = ""
        }
        
        # Get admin email for ownership assignments
        $script:TenantState.AdminEmail = Read-Host "Enter the email address for the Global Admin account"
        
        return $true
    }
    catch {
        Write-LogMessage -Message "Error verifying tenant domain - $($_.Exception.Message)" -Type Error
        return $false
    }
}

# === Group Creation Functions ===
function New-TenantGroups {
    Write-LogMessage -Message "Starting group creation process..." -Type Info
    Import-RequiredGraphModules
    
    try {
        $tenantName = $script:TenantState.TenantName
        Write-LogMessage -Message "Creating groups for tenant: $tenantName" -Type Info
        
        # Create license groups
        foreach ($license in $config.DefaultGroups.License) {
            $displayName = "Microsoft 365 $license Users"
            
            # Check if already exists
            $existingGroup = Get-MgGroup -Filter "displayName eq '$displayName'" -ErrorAction SilentlyContinue
            if ($existingGroup) {
                Write-LogMessage -Message "Group '$displayName' already exists" -Type Warning
                $script:TenantState.CreatedGroups[$displayName] = $existingGroup.Id
                continue
            }
            
            # Create using direct API
            $body = @{
                displayName = $displayName
                description = "Dynamic license group for $license"
                groupTypes = @("DynamicMembership")
                mailEnabled = $false
                mailNickname = "$($license)Users"
                membershipRule = "user.extensionAttribute1 eq `"$license`""
                membershipRuleProcessingState = "On"
                securityEnabled = $true
            }

            try {
                $result = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/groups" -Body $body
                Write-LogMessage -Message "Created dynamic license group: $displayName" -Type Success
                $script:TenantState.CreatedGroups[$displayName] = $result.id
            }
            catch {
                Write-LogMessage -Message "Failed to create $displayName - $($_.Exception.Message)" -Type Error
            }
        }

        # Create domain users group
        if (-not [string]::IsNullOrEmpty($tenantName)) {
            $domainGroupName = "$tenantName Users"
            # Check if already exists
            $existingGroup = Get-MgGroup -Filter "displayName eq '$domainGroupName'" -ErrorAction SilentlyContinue
            
            if (-not $existingGroup) {
                $body = @{
                    displayName = $domainGroupName
                    description = "All users in $tenantName tenant"
                    groupTypes = @("DynamicMembership")
                    mailEnabled = $false
                    mailNickname = "DomainUsers"
                    membershipRule = "user.userType -ne `"Guest`""
                    membershipRuleProcessingState = "On"
                    securityEnabled = $true
                }

                try {
                    $result = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/groups" -Body $body
                    Write-LogMessage -Message "Created domain users group: $domainGroupName" -Type Success
                    $script:TenantState.CreatedGroups[$domainGroupName] = $result.id
                }
                catch {
                    Write-LogMessage -Message "Failed to create domain users group - $($_.Exception.Message)" -Type Error
                }
            }
            else {
                Write-LogMessage -Message "Domain users group already exists" -Type Warning
                $script:TenantState.CreatedGroups[$domainGroupName] = $existingGroup.Id
            }
        }

        # Create regular security groups
        foreach ($name in $config.DefaultGroups.Security) {
            # Check if already exists
            $existingGroup = Get-MsgGroup -Filter "displayName eq '$name'" -ErrorAction SilentlyContinue
            if ($existingGroup) {
                Write-LogMessage -Message "Group '$name' already exists" -Type Warning
                $script:TenantState.CreatedGroups[$name] = $existingGroup.Id
                continue
            }
            
            $mailNick = $name -replace '\s', ''
            
            try {
                $newGroup = New-MgGroup -DisplayName $name -Description "Security group" -MailEnabled:$false -MailNickname $mailNick -SecurityEnabled:$true
                Write-LogMessage -Message "Created security group: $name" -Type Success
                $script:TenantState.CreatedGroups[$name] = $newGroup.Id
            }
            catch {
                Write-LogMessage -Message "Failed to create group $name - $($_.Exception.Message)" -Type Error
            }
        }

        Write-LogMessage -Message "Group creation completed" -Type Success
        return $true
    }
    catch {
        Write-LogMessage -Message "Error in group creation process - $($_.Exception.Message)" -Type Error
        return $false
    }
}

# === Conditional Access Policy Functions ===
function New-TenantCAPolices {
    Write-LogMessage -Message "Starting CA policy creation process..." -Type Info
    Import-RequiredGraphModules
    
    try {
        # Check for NoMFA Exemption group ID
        $noMfaGroupId = $script:TenantState.CreatedGroups["NoMFA Exemption"]
        if (-not $noMfaGroupId) {
            Write-LogMessage -Message "NoMFA Exemption group not found. Some policies may not be correctly configured." -Type Warning
        }
        
        # Function to check if policy exists using direct API
        function Test-PolicyExists {
            param ([string]$PolicyName)
            
            try {
                $response = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies" -ErrorAction Stop
                
                if ($response.PSObject.Properties.Name -contains "value") {
                    $policies = $response.value
                } else {
                    $policies = @($response)
                }
                
                foreach ($p in $policies) {
                    if ($p.displayName -eq $PolicyName) {
                        return $true
                    }
                }
                return $false
            }
            catch {
                Write-LogMessage -Message "Error checking policies - $($_.Exception.Message)" -Type Error
                return $false
            }
        }

        # Create C001 - Block Legacy Authentication
        $policyName = "C001 - Block Legacy Authentication All Apps"
        if (Test-PolicyExists -PolicyName $policyName) {
            Write-LogMessage -Message "Policy '$policyName' already exists" -Type Warning
        }
        else {
            $body = @{
                displayName = $policyName
                state = "disabled"
                conditions = @{
                    users = @{
                        includeUsers = @("All")
                    }
                    applications = @{
                        includeApplications = @("All")
                    }
                    clientAppTypes = @("exchangeActiveSync", "other")
                }
                grantControls = @{
                    operator = "OR"
                    builtInControls = @("block")
                }
            }
            
            try {
                Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies" -Body $body -ErrorAction Stop
                Write-LogMessage -Message "Created policy: $policyName" -Type Success
            }
            catch {
                Write-LogMessage -Message "Failed to create policy $policyName - $($_.Exception.Message)" -Type Error
            }
        }

        # Create C002 - MFA Required for All Users
        $policyName = "C002 - MFA Required for All Users"
        if (Test-PolicyExists -PolicyName $policyName) {
            Write-LogMessage -Message "Policy '$policyName' already exists" -Type Warning
        }
        else {
            $body = @{
                displayName = $policyName
                state = "enabled"
                conditions = @{
                    users = @{
                        includeUsers = @("All")
                    }
                    applications = @{
                        includeApplications = @("All")
                    }
                    clientAppTypes = @("browser", "mobileAppsAndDesktopClients")
                }
                grantControls = @{
                    operator = "OR"
                    builtInControls = @("mfa")
                }
            }
            
            # Add NoMFA group exclusion if available
            if ($noMfaGroupId) {
                $body.conditions.users.excludeGroups = @($noMfaGroupId)
                Write-LogMessage -Message "Added NoMFA Exemption group to exclusions" -Type Info
            }
            
            try {
                Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies" -Body $body -ErrorAction Stop
                Write-LogMessage -Message "Created policy: $policyName" -Type Success
            }
            catch {
                Write-LogMessage -Message "Failed to create policy $policyName - $($_.Exception.Message)" -Type Error
            }
        }

        # Create C003 - Block Non Corporate Devices
        $policyName = "C003 - Block Non Corporate Devices"
        if (Test-PolicyExists -PolicyName $policyName) {
            Write-LogMessage -Message "Policy '$policyName' already exists" -Type Warning
        }
        else {
            $body = @{
                displayName = $policyName
                state = "enabledForReportingButNotEnforced"
                conditions = @{
                    users = @{
                        includeUsers = @("All")
                        excludeRoles = @("d29b2b05-8046-44ba-8758-1e26182fcf32")  # Global Admin role
                    }
                    applications = @{
                        includeApplications = @("All")
                    }
                    clientAppTypes = @("all")
                    platforms = @{
                        includePlatforms = @("all")
                    }
                }
                grantControls = @{
                    operator = "OR"
                    builtInControls = @("compliantDevice", "domainJoinedDevice")
                }
            }
            
            try {
                Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies" -Body $body -ErrorAction Stop
                Write-LogMessage -Message "Created policy: $policyName" -Type Success
            }
            catch {
                Write-LogMessage -Message "Failed to create policy $policyName - $($_.Exception.Message)" -Type Error
            }
        }

        # Create C004 - Require Password Change and MFA for High Risk Users
        $policyName = "C004 - Require Password Change and MFA for High Risk Users"
        if (Test-PolicyExists -PolicyName $policyName) {
            Write-LogMessage -Message "Policy '$policyName' already exists" -Type Warning
        }
        else {
            $body = @{
                displayName = $policyName
                state = "enabled"
                conditions = @{
                    users = @{
                        includeUsers = @("All")
                    }
                    applications = @{
                        includeApplications = @("All")
                    }
                    userRiskLevels = @("high")
                }
                grantControls = @{
                    operator = "AND"
                    builtInControls = @("mfa", "passwordChange")
                }
            }
            
            try {
                Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies" -Body $body -ErrorAction Stop
                Write-LogMessage -Message "Created policy: $policyName" -Type Success
            }
            catch {
                Write-LogMessage -Message "Failed to create policy $policyName - $($_.Exception.Message)" -Type Error
            }
        }

        # Create C005 - Require MFA for Risky Sign-Ins
        $policyName = "C005 - Require MFA for Risky Sign-Ins"
        if (Test-PolicyExists -PolicyName $policyName) {
            Write-LogMessage -Message "Policy '$policyName' already exists" -Type Warning
        }
        else {
            $body = @{
                displayName = $policyName
                state = "enabled"
                conditions = @{
                    users = @{
                        includeUsers = @("All")
                    }
                    applications = @{
                        includeApplications = @("All")
                    }
                    signInRiskLevels = @("high", "medium")
                }
                grantControls = @{
                    operator = "OR"
                    builtInControls = @("mfa")
                }
            }
            
            try {
                Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies" -Body $body -ErrorAction Stop
                Write-LogMessage -Message "Created policy: $policyName" -Type Success
            }
            catch {
                Write-LogMessage -Message "Failed to create policy $policyName - $($_.Exception.Message)" -Type Error
            }
        }

        Write-LogMessage -Message "CA Policy setup completed" -Type Success
        return $true
    }
    catch {
        Write-LogMessage -Message "Error in CA policy creation process - $($_.Exception.Message)" -Type Error
        return $false
    }
}

# === SharePoint Configuration Functions ===
function New-TenantSharePoint {
    Write-LogMessage -Message "Starting SharePoint configuration..." -Type Info
    
    # Clear SharePoint authentication cache to prevent conflicts
    Write-LogMessage -Message "Clearing SharePoint authentication cache..." -Type Info
    try {
        Disconnect-MsgGraph -ErrorAction SilentlyContinue | Out-Null
        Disconnect-SPOService -ErrorAction SilentlyContinue | Out-Null
        Remove-Item "$env:USERPROFILE\.mg" -Recurse -Force -ErrorAction SilentlyContinue
    }
    catch {
        # Ignore cleanup errors
    }
    
    # Reconnect to Microsoft Graph (needed for security group creation)
    Write-LogMessage -Message "Reconnecting to Microsoft Graph..." -Type Info
    try {
        Import-RequiredGraphModules
        Connect-MgGraph -Scopes $config.GraphScopes -ErrorAction Stop | Out-Null
        $context = Get-MgContext
        Write-LogMessage -Message "Successfully reconnected to Microsoft Graph as $($context.Account)" -Type Success
    }
    catch {
        Write-LogMessage -Message "Failed to reconnect to Microsoft Graph - $($_.Exception.Message)" -Type Error
        return $false
    }
    
    try {
        # Get SharePoint URLs - simplified input
        $customerName = $script:TenantState.TenantName
        Write-Host "SharePoint URL Configuration" -ForegroundColor Yellow
        Write-Host "Example: If your tenant is 'm365x36060197.sharepoint.com', enter 'm365x36060197'" -ForegroundColor Cyan
        $tenantName = Read-Host "Enter your SharePoint tenant name (without .sharepoint.com)"
        
        # Construct URLs automatically
        $adminUrl = "https://$tenantName-admin.sharepoint.com"
        $tenantUrl = "https://$tenantName.sharepoint.com"
        $ownerEmail = $script:TenantState.AdminEmail
        
        Write-LogMessage -Message "SharePoint Admin URL: $adminUrl" -Type Info
        Write-LogMessage -Message "SharePoint Tenant URL: $tenantUrl" -Type Info
        
        # Connect to SharePoint Online Admin Center
        Write-LogMessage -Message "Connecting to SharePoint Online Admin Center..." -Type Info
        Connect-SPOService -Url $adminUrl
        
        # Create a Hub Site
        $hubSiteTitle = "$customerName Hub"
        $hubSiteUrl = "$tenantUrl/sites/corporatehub"
        
        # Check if the Hub Site already exists, if not, create one
        try {
            $existingHubSite = Get-SPOSite | Where-Object { $_.Url -eq $hubSiteUrl }
            if ($existingHubSite) {
                Write-LogMessage -Message "Hub site already exists: $hubSiteUrl" -Type Warning
            } else {
                New-SPOSite -Url $hubSiteUrl -Owner $ownerEmail -StorageQuota $config.SharePoint.StorageQuota -Title $hubSiteTitle -Template $config.SharePoint.SiteTemplate
                Write-LogMessage -Message "Hub site created: $hubSiteUrl" -Type Success
            }
        }
        catch {
            Write-LogMessage -Message "Hub site may already exist or creation failed: $($_.Exception.Message)" -Type Warning
        }
        
        # Set the site as a Hub Site
        try {
            $principals = @($ownerEmail)
            Register-SPOHubSite -Site $hubSiteUrl -Principals $principals -ErrorAction SilentlyContinue
            Write-LogMessage -Message "Hub site registered: $hubSiteUrl" -Type Success
        }
        catch {
            Write-LogMessage -Message "Hub site registration failed (may already be registered): $($_.Exception.Message)" -Type Warning
        }
        
        # Create spokes sites array
        $spokeSites = @()
        foreach ($siteName in $config.SharePoint.DefaultSites) {
            $spokeSites += @{ 
                Name = $siteName
                URL = "$tenantUrl/sites/$($siteName.ToLower())" 
            }
        }
        
        # Create security groups for each site
        $securityGroups = @{}
        Import-RequiredGraphModules
        
        foreach ($site in $spokeSites) {
            $siteName = $site.Name
            Write-LogMessage -Message "Creating security groups for site: $siteName" -Type Info
            
            foreach ($groupType in @("Members", "Owners", "Visitors")) {
                $groupName = "$siteName SharePoint $groupType"
                $existingGroup = Get-MgGroup -Filter "DisplayName eq '$groupName'" -ErrorAction SilentlyContinue
                
                if ($existingGroup -eq $null) {
                    try {
                        $newGroup = New-MgGroup -DisplayName $groupName -MailEnabled:$false -MailNickname "$siteName-SPO-$groupType" -SecurityEnabled:$true
                        $securityGroups["$siteName-$groupType"] = $newGroup.Id
                        Write-LogMessage -Message "Security group created: $groupName" -Type Success
                    } catch {
                        Write-LogMessage -Message "Failed to create security group: $groupName - $($_.Exception.Message)" -Type Error
                        continue
                    }
                } else {
                    $securityGroups["$siteName-$groupType"] = $existingGroup.Id
                    Write-LogMessage -Message "Security group already exists: $groupName" -Type Warning
                }
            }
        }
        
        # Wait for security groups to propagate
        Write-LogMessage -Message "Waiting for security groups to propagate (2 minutes)..." -Type Info
        for ($i = 0; $i -lt 120; $i++) {
            Start-Sleep -Seconds 1
            Show-Progress -Current ($i + 1) -Total 120 -Status "Waiting for groups to propagate..."
        }
        Write-Host ""
        
        # Create spoke sites
        foreach ($site in $spokeSites) {
            $siteUrl = $site.URL
            $siteName = $site.Name
            
            # Check if the Spoke Site already exists, if not, create one
            try {
                $existingSpokeSite = Get-SPOSite | Where-Object { $_.Url -eq $siteUrl }
                if ($existingSpokeSite) {
                    Write-LogMessage -Message "$siteName site already exists: $siteUrl" -Type Warning
                } else {
                    New-SPOSite -Url $siteUrl -Owner $ownerEmail -StorageQuota $config.SharePoint.StorageQuota -Title "$siteName" -Template $config.SharePoint.SiteTemplate
                    Write-LogMessage -Message "$siteName site created: $siteUrl" -Type Success
                }
            }
            catch {
                # If checking fails, try to create anyway - might be a permissions issue with Get-SPOSite
                try {
                    New-SPOSite -Url $siteUrl -Owner $ownerEmail -StorageQuota $config.SharePoint.StorageQuota -Title "$siteName" -Template $config.SharePoint.SiteTemplate
                    Write-LogMessage -Message "$siteName site created: $siteUrl" -Type Success
                }
                catch {
                    if ($_.Exception.Message -like "*already exists*" -or $_.Exception.Message -like "*site collection*already*") {
                        Write-LogMessage -Message "$siteName site already exists: $siteUrl" -Type Warning
                    } else {
                        Write-LogMessage -Message "Failed to create $siteName site: $($_.Exception.Message)" -Type Error
                        continue
                    }
                }
            }
        
            # Register Spoke Site to the Hub
            try {
                Add-SPOHubSiteAssociation -Site $siteUrl -HubSite $hubSiteUrl -ErrorAction SilentlyContinue
                Write-LogMessage -Message "$siteName site associated with Hub site" -Type Success
            }
            catch {
                Write-LogMessage -Message "Failed to associate $siteName site with hub (may already be associated): $($_.Exception.Message)" -Type Warning
            }
        }
        
        # Wait for sites to provision
        Write-LogMessage -Message "Waiting for SharePoint sites to provision (3 minutes)..." -Type Info
        for ($i = 0; $i -lt 180; $i++) {
            Start-Sleep -Seconds 1
            Show-Progress -Current ($i + 1) -Total 180 -Status "Waiting for sites to provision..."
        }
        Write-Host ""
        
        # Add security groups to sites
        foreach ($site in $spokeSites) {
            $siteUrl = $site.URL
            $siteName = $site.Name
            
            foreach ($groupType in @("Members", "Owners", "Visitors")) {
                $groupKey = "$siteName-$groupType"
                $spoGroupName = "$siteName $groupType"
                
                if ($securityGroups.ContainsKey($groupKey)) {
                    $groupId = $securityGroups[$groupKey]
                    try {
                        Add-SPOUser -Site $siteUrl -Group $spoGroupName -LoginName "c:0t.c|tenant|$groupId" -ErrorAction Stop
                        Write-LogMessage -Message "Added security group as $groupType to $siteUrl" -Type Success
                    } catch {
                        # Check if it's already added
                        if ($_.Exception.Message -like "*already exists*" -or $_.Exception.Message -like "*already a member*") {
                            Write-LogMessage -Message "Security group already exists as $groupType in $siteUrl" -Type Warning
                        } else {
                            Write-LogMessage -Message "Failed to add security group to $siteUrl - $($_.Exception.Message)" -Type Error
                        }
                    }
                } else {
                    Write-LogMessage -Message "No security group found for $siteName $groupType. Skipping..." -Type Warning
                }
            }
        }
        
        Write-LogMessage -Message "SharePoint configuration completed" -Type Success
        Disconnect-SPOService
        return $true
    }
    catch {
        Write-LogMessage -Message "Error in SharePoint configuration - $($_.Exception.Message)" -Type Error
        return $false
    }
}

# === Intune Configuration Functions ===
function New-TenantIntune {
    Write-LogMessage -Message "Starting Intune configuration..." -Type Info
    Import-RequiredGraphModules
    
    try {
        # Placeholder for Intune configuration
        # Future implementation will include:
        # - Device compliance policies
        # - Device configuration profiles  
        # - App protection policies
        # - Enrollment restrictions
        # - Windows Autopilot configuration
        
        Write-LogMessage -Message "Intune configuration not yet implemented - focusing on core components first" -Type Warning
        Write-LogMessage -Message "This will be implemented in a future version" -Type Info
        return $true
    }
    catch {
        Write-LogMessage -Message "Error in Intune configuration - $($_.Exception.Message)" -Type Error
        return $false
    }
}

# === Helper Functions for User Creation ===
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

# === Helper Functions for User Creation (from working script) ===
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

function Get-ExcelFile {
    param (
        [string]$DefaultPath
    )
    
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
        
        # Debug the password data first
        Write-LogMessage -Message "Raw password value for $($user.DisplayName): '$($user.Password)'" -Type Info -LogOnly
        Write-LogMessage -Message "Password type: $($user.Password.GetType().Name)" -Type Info -LogOnly
        Write-LogMessage -Message "Password is null: $($user.Password -eq $null)" -Type Info -LogOnly
        Write-LogMessage -Message "Password is empty string: $([string]::IsNullOrEmpty($user.Password))" -Type Info -LogOnly
        Write-LogMessage -Message "Password is whitespace: $([string]::IsNullOrWhiteSpace($user.Password))" -Type Info -LogOnly
        
        # Skip if required fields are missing
        if (-not (Test-NotEmpty -Value $user.UserPrincipalName) -or 
            -not (Test-NotEmpty -Value $user.DisplayName) -or
            -not (Test-NotEmpty -Value $user.Password)) {
            
            Write-LogMessage -Message "Skipping user with missing required fields: $($user.DisplayName)" -Type Warning -LogOnly
            $results.Skipped += $user.DisplayName
            continue
        }
        
        # Additional password validation
        if ([string]::IsNullOrWhiteSpace([string]$user.Password)) {
            Write-LogMessage -Message "Skipping user $($user.DisplayName) - password is null or whitespace" -Type Warning -LogOnly
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
        
        # Create password profile - EXACT syntax from working script
        $passwordProfile = @{
            Password = $user.Password
            ForceChangePasswordNextSignIn = $true
        }
        
        # Create user parameters - EXACT syntax from working script
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
        
        # Create user - EXACT syntax from working script
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

function New-TenantUsers {
    Write-LogMessage -Message "Starting user creation process..." -Type Info
    Import-RequiredGraphModules
    
    try {
        # Get user data source
        Add-Type -AssemblyName System.Windows.Forms
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.Title = "Select Users Excel File"
        $openFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx"
        
        if ($openFileDialog.ShowDialog() -eq 'OK') {
            $excelPath = $openFileDialog.FileName
            
            # Import user data - using working approach
            $users = Import-Excel -Path $excelPath -WorksheetName "in" | Where-Object { 
                (Test-NotEmpty -Value $_.UserPrincipalName) -and
                (Test-NotEmpty -Value $_.DisplayName) -and
                (Test-NotEmpty -Value $_.Password)
            }
            
            if ($users.Count -eq 0) {
                Write-LogMessage -Message "No valid users found in Excel file" -Type Error
                return $false
            }
            
            # Create users - EXACT logic from working script
            $totalUsers = $users.Count
            $currentUser = 0
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
            
            foreach ($user in $users) {
                $currentUser++
                # FIXED: Added space after colon to prevent parser error
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
            
            # Display results
            Write-LogMessage -Message "User creation completed" -Type Success
            Write-LogMessage -Message "Total users created: $($results.Success.Count)" -Type Info
            Write-LogMessage -Message "Total users failed: $($results.Failed.Count)" -Type Info
            Write-LogMessage -Message "Total users skipped: $($results.Skipped.Count)" -Type Info
            Write-LogMessage -Message "Total managers set: $($results.ManagersSet.Count)" -Type Info
            Write-LogMessage -Message "Total license attributes set: $($results.LicensesSet.Count)" -Type Info
            
            return $true
        }
        else {
            Write-LogMessage -Message "User file selection canceled" -Type Warning
            return $false
        }
    }
    catch {
        Write-LogMessage -Message "Error in user creation process - $($_.Exception.Message)" -Type Error
        return $false
    }
}

# === Documentation Generation Functions ===
function New-TenantDocumentation {
    Write-LogMessage -Message "Starting documentation generation..." -Type Info
    
    try {
        # Placeholder for documentation generation
        # This would be implemented with actual documentation generation
        # outputting to Word or Excel
        
        Write-LogMessage -Message "Documentation generation not yet implemented" -Type Warning
        return $true
    }
    catch {
        Write-LogMessage -Message "Error in documentation generation - $($_.Exception.Message)" -Type Error
        return $false
    }
}

# === Main Execution ===
function Start-Setup {
    # Clear any existing authentication state to prevent conflicts
    Write-Host "Clearing authentication cache..." -ForegroundColor Cyan
    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        Disconnect-SPOService -ErrorAction SilentlyContinue | Out-Null
        Remove-Item "$env:USERPROFILE\.mg" -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "Authentication cache cleared" -ForegroundColor Green
    }
    catch {
        # Ignore cleanup errors
    }
    
    # Initialize log file
    try {
        $logDir = [System.IO.Path]::GetDirectoryName($config.LogFile)
        if (-not (Test-Path -Path $logDir)) {
            New-Item -Path $logDir -ItemType Directory -Force | Out-Null
        }
        
        New-Item -Path $config.LogFile -ItemType File -Force | Out-Null
        Write-LogMessage -Message "Unified Microsoft 365 Tenant Setup Utility started" -Type Info
        Write-LogMessage -Message "PowerShell version: $($PSVersionTable.PSVersion)" -Type Info -LogOnly
        Write-LogMessage -Message "Computer name: $env:COMPUTERNAME" -Type Info -LogOnly
        Write-LogMessage -Message "User context: $env:USERNAME" -Type Info -LogOnly
    }
    catch {
        Write-Host "Failed to start logging: $_" -ForegroundColor Red
        exit
    }
    
    # Check execution policy
    Write-LogMessage -Message "Checking execution policy..." -Type Info
    $currentPolicy = Get-ExecutionPolicy
    if ($currentPolicy -eq 'Restricted') {
        Write-LogMessage -Message "PowerShell execution policy is set to Restricted. This may prevent the script from running properly." -Type Warning
        Write-Host "Consider running: Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser" -ForegroundColor Yellow
        $continue = Read-Host "Do you want to continue anyway? (Y/N)"
        if ($continue -ne 'Y' -and $continue -ne 'y') {
            Write-LogMessage -Message "Script execution cancelled due to execution policy." -Type Info
            exit
        }
    }
    
    # Check and install required modules
    $modulesInstalled = Install-RequiredModules
    if (-not $modulesInstalled) {
        Write-LogMessage -Message "Required modules installation failed. Exiting." -Type Error
        Read-Host "Press Enter to exit"
        exit
    }
    
    # Main menu loop
    $exitScript = $false
    while (-not $exitScript) {
        $choice = Show-Menu -Title "Main Menu" -Options @(
            "Connect to Microsoft Graph and Verify Tenant"
            "Create Security and License Groups"
            "Configure Conditional Access Policies"
            "Set Up SharePoint Sites"
            "Configure Intune Policies"
            "Create Users"
            "Generate Documentation"
            "Debug Excel File (Check Password Data)"
            "Exit"
        )
        
        switch ($choice) {
            1 {
                # Connect to Graph and verify tenant
                $connected = Connect-ToGraphAndVerify
                if ($connected) {
                    Write-LogMessage -Message "Successfully connected and verified tenant domain" -Type Success
                }
                Read-Host "Press Enter to continue"
            }
            2 {
                # Create groups
                if (-not (Get-MgContext)) {
                    Write-LogMessage -Message "Not connected to Microsoft Graph. Please connect first." -Type Warning
                }
                else {
                    $groupsCreated = New-TenantGroups
                }
                Read-Host "Press Enter to continue"
            }
            3 {
                # Configure CA policies
                if (-not (Get-MgContext)) {
                    Write-LogMessage -Message "Not connected to Microsoft Graph. Please connect first." -Type Warning
                }
                elseif (-not $script:TenantState -or -not $script:TenantState.CreatedGroups) {
                    Write-LogMessage -Message "Groups not created yet. Please create groups first." -Type Warning
                }
                else {
                    $policiesCreated = New-TenantCAPolices
                }
                Read-Host "Press Enter to continue"
            }
            4 {
                # Set up SharePoint
                if (-not (Get-MgContext)) {
                    Write-LogMessage -Message "Not connected to Microsoft Graph. Please connect first." -Type Warning
                }
                else {
                    $sharePointSetup = New-TenantSharePoint
                }
                Read-Host "Press Enter to continue"
            }
            5 {
                # Configure Intune
                if (-not (Get-MgContext)) {
                    Write-LogMessage -Message "Not connected to Microsoft Graph. Please connect first." -Type Warning
                }
                else {
                    $intuneSetup = New-TenantIntune
                }
                Read-Host "Press Enter to continue"
            }
            6 {
                # Create users
                if (-not (Get-MgContext)) {
                    Write-LogMessage -Message "Not connected to Microsoft Graph. Please connect first." -Type Warning
                }
                else {
                    $usersCreated = New-TenantUsers
                }
                Read-Host "Press Enter to continue"
            }
            7 {
                # Generate documentation
                if (-not $script:TenantState) {
                    Write-LogMessage -Message "No tenant configuration found. Please connect and configure tenant first." -Type Warning
                }
                else {
                    $docGenerated = New-TenantDocumentation
                }
                Read-Host "Press Enter to continue"
            }
            8 {
                # Debug Excel file
                Debug-ExcelData
                Read-Host "Press Enter to continue"
            }
            9 {
                # Exit
                $exitScript = $true
                Write-LogMessage -Message "Unified Microsoft 365 Tenant Setup Utility ended" -Type Info
            }
        }
    }

    # Final cleanup
    if (Get-MgContext) {
        Disconnect-MgGraph | Out-Null
    }

    Write-Host ""
    Write-Host "Thank you for using the Unified Microsoft 365 Tenant Setup Utility!" -ForegroundColor Cyan
}

# Start the setup
Start-Setup