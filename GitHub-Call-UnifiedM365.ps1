#requires -Version 5.1
<#
.SYNOPSIS
    Main Setup Script for Microsoft 365 Tenant Setup Utility
.DESCRIPTION
    Provides core infrastructure, menu interface, and Graph connection for Microsoft 365 tenant setup operations
.NOTES
    Version: 1.0
    Requirements: PowerShell 5.1 or later, Microsoft Graph PowerShell SDK
#>

# === Configuration ===
$config = @{
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
}

# === Core Helper Functions ===
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

# === Graph Connection Functions ===
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

# === GitHub Module Loading Functions ===
$GitHubConfig = @{
    BaseUrl = "https://raw.githubusercontent.com/iceedd/BITS-Projects/main"
    CacheDirectory = "$env:TEMP\M365TenantSetup\Modules"
    ModuleFiles = @{
        "Groups" = "Groups-Module.ps1"
        "ConditionalAccess" = "Conditional-Access-Module.ps1"
        "SharePoint" = "SharePoint-Module.ps1"
        "Intune" = "Intune-Module.ps1"
        "Users" = "User-Module.ps1"
        "Documentation" = "Documentation-Module.ps1"
    }
}

function Initialize-ModuleCache {
    try {
        if (-not (Test-Path -Path $GitHubConfig.CacheDirectory)) {
            New-Item -Path $GitHubConfig.CacheDirectory -ItemType Directory -Force | Out-Null
            Write-LogMessage -Message "Created module cache directory: $($GitHubConfig.CacheDirectory)" -Type Info -LogOnly
        }
        return $true
    }
    catch {
        Write-LogMessage -Message "Failed to create module cache directory: $($_.Exception.Message)" -Type Error
        return $false
    }
}

function Get-ModuleFromGitHub {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ModuleName
    )
    
    try {
        if (-not $GitHubConfig.ModuleFiles.ContainsKey($ModuleName)) {
            Write-LogMessage -Message "Unknown module: $ModuleName" -Type Error
            return $false
        }
        
        $fileName = $GitHubConfig.ModuleFiles[$ModuleName]
        $moduleUrl = "$($GitHubConfig.BaseUrl)/$fileName"
        $localPath = Join-Path -Path $GitHubConfig.CacheDirectory -ChildPath $fileName
        
        Write-LogMessage -Message "Downloading $ModuleName module from GitHub..." -Type Info
        Write-LogMessage -Message "URL: $moduleUrl" -Type Info -LogOnly
        
        # Download the module
        Invoke-WebRequest -Uri $moduleUrl -OutFile $localPath -ErrorAction Stop
        
        Write-LogMessage -Message "Successfully downloaded $ModuleName module" -Type Success -LogOnly
        return $localPath
    }
    catch {
        Write-LogMessage -Message "Failed to download $ModuleName module: $($_.Exception.Message)" -Type Error
        return $false
    }
}

function Import-ModuleFromCache {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ModuleName,
        
        [Parameter(Mandatory = $false)]
        [switch]$ForceRefresh
    )
    
    try {
        $fileName = $GitHubConfig.ModuleFiles[$ModuleName]
        $localPath = Join-Path -Path $GitHubConfig.CacheDirectory -ChildPath $fileName
        
        # Check if module exists locally and if we should refresh
        $shouldDownload = $ForceRefresh -or (-not (Test-Path -Path $localPath))
        
        if ($shouldDownload) {
            $downloadResult = Get-ModuleFromGitHub -ModuleName $ModuleName
            if (-not $downloadResult) {
                return $false
            }
            $localPath = $downloadResult
        }
        else {
            Write-LogMessage -Message "Using cached $ModuleName module" -Type Info -LogOnly
        }
        
        Write-LogMessage -Message "Loading $ModuleName module from: $localPath" -Type Info
        
        # Check if file exists and has content
        if (-not (Test-Path -Path $localPath)) {
            Write-LogMessage -Message "Module file not found at: $localPath" -Type Error
            return $false
        }
        
        $fileSize = (Get-Item $localPath).Length
        Write-LogMessage -Message "Module file size: $fileSize bytes" -Type Info -LogOnly
        
        if ($fileSize -eq 0) {
            Write-LogMessage -Message "Module file is empty" -Type Error
            return $false
        }
        
        # Execute the module file directly in the global scope using dot sourcing with proper path
        . $localPath
        
        Write-LogMessage -Message "Successfully loaded $ModuleName module" -Type Success
        return $true
    }
    catch {
        Write-LogMessage -Message "Failed to load $ModuleName module: $($_.Exception.Message)" -Type Error
        Write-LogMessage -Message "Error details: $($_.Exception.ToString())" -Type Error -LogOnly
        return $false
    }
}

function Test-ModuleAvailability {
    param (
        [Parameter(Mandatory = $true)]
        [string]$ModuleName
    )
    
    try {
        $fileName = $GitHubConfig.ModuleFiles[$ModuleName]
        $moduleUrl = "$($GitHubConfig.BaseUrl)/$fileName"
        
        # Test if the module exists on GitHub
        $response = Invoke-WebRequest -Uri $moduleUrl -Method Head -ErrorAction Stop
        return $response.StatusCode -eq 200
    }
    catch {
        Write-LogMessage -Message "Module $ModuleName not available on GitHub: $($_.Exception.Message)" -Type Warning -LogOnly
        return $false
    }
}

# === Menu Functions ===
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
    
    # Initialize module cache
    $cacheInitialized = Initialize-ModuleCache
    if (-not $cacheInitialized) {
        Write-LogMessage -Message "Failed to initialize module cache. Some features may not work." -Type Warning
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
            "Refresh Modules from GitHub"
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
                    $moduleLoaded = Import-ModuleFromCache -ModuleName "Groups"
                    if ($moduleLoaded) {
                        $groupsCreated = New-TenantGroups
                    }
                    else {
                        Write-LogMessage -Message "Failed to load Groups module. Please check your internet connection." -Type Error
                    }
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
                    $moduleLoaded = Import-ModuleFromCache -ModuleName "ConditionalAccess"
                    if ($moduleLoaded) {
                        $policiesCreated = New-TenantCAPolices
                    }
                    else {
                        Write-LogMessage -Message "Failed to load Conditional Access module. Please check your internet connection." -Type Error
                    }
                }
                Read-Host "Press Enter to continue"
            }
            4 {
                # Set up SharePoint
                if (-not (Get-MgContext)) {
                    Write-LogMessage -Message "Not connected to Microsoft Graph. Please connect first." -Type Warning
                }
                else {
                    $moduleLoaded = Import-ModuleFromCache -ModuleName "SharePoint"
                    if ($moduleLoaded) {
                        $sharePointSetup = New-TenantSharePoint
                    }
                    else {
                        Write-LogMessage -Message "Failed to load SharePoint module. Please check your internet connection." -Type Error
                    }
                }
                Read-Host "Press Enter to continue"
            }
            5 {
                # Configure Intune
                if (-not (Get-MgContext)) {
                    Write-LogMessage -Message "Not connected to Microsoft Graph. Please connect first." -Type Warning
                }
                else {
                    $moduleLoaded = Import-ModuleFromCache -ModuleName "Intune"
                    if ($moduleLoaded) {
                        $intuneSetup = New-TenantIntune
                    }
                    else {
                        Write-LogMessage -Message "Failed to load Intune module. Please check your internet connection." -Type Error
                    }
                }
                Read-Host "Press Enter to continue"
            }
            6 {
                # Create users
                if (-not (Get-MgContext)) {
                    Write-LogMessage -Message "Not connected to Microsoft Graph. Please connect first." -Type Warning
                }
                else {
                    $moduleLoaded = Import-ModuleFromCache -ModuleName "Users"
                    if ($moduleLoaded) {
                        Write-LogMessage -Message "Attempting to call New-TenantUsers..." -Type Info
                        
                        # Test if the function is available by trying to get its definition
                        try {
                            $functionDef = Get-Command New-TenantUsers -ErrorAction Stop
                            Write-LogMessage -Message "New-TenantUsers function found with source: $($functionDef.Source)" -Type Success
                            
                            # Call the function
                            $usersCreated = New-TenantUsers
                            Write-LogMessage -Message "New-TenantUsers completed successfully" -Type Success
                        }
                        catch [System.Management.Automation.CommandNotFoundException] {
                            Write-LogMessage -Message "New-TenantUsers function not found after module load" -Type Error
                            
                            # Try alternative function calling methods
                            Write-LogMessage -Message "Attempting direct execution from module file..." -Type Info
                            try {
                                # Get the cached module path and execute it directly
                                $fileName = $GitHubConfig.ModuleFiles["Users"]
                                $localPath = Join-Path -Path $GitHubConfig.CacheDirectory -ChildPath $fileName
                                
                                # Source the file and call the function in one go
                                $result = & {
                                    . $localPath
                                    New-TenantUsers
                                }
                                Write-LogMessage -Message "Direct execution successful" -Type Success
                            }
                            catch {
                                Write-LogMessage -Message "Direct execution also failed: $($_.Exception.Message)" -Type Error
                                Write-LogMessage -Message "This indicates a dependency issue with core functions" -Type Warning
                            }
                        }
                        catch {
                            Write-LogMessage -Message "Error calling New-TenantUsers: $($_.Exception.Message)" -Type Error
                        }
                    }
                    else {
                        Write-LogMessage -Message "Failed to load Users module. Please check your internet connection." -Type Error
                    }
                }
                Read-Host "Press Enter to continue"
            }
            7 {
                # Generate documentation
                if (-not $script:TenantState) {
                    Write-LogMessage -Message "No tenant configuration found. Please connect and configure tenant first." -Type Warning
                }
                else {
                    $moduleLoaded = Import-ModuleFromCache -ModuleName "Documentation"
                    if ($moduleLoaded) {
                        $docGenerated = New-TenantDocumentation
                    }
                    else {
                        Write-LogMessage -Message "Failed to load Documentation module. Please check your internet connection." -Type Error
                    }
                }
                Read-Host "Press Enter to continue"
            }
            8 {
                # Debug Excel file
                $moduleLoaded = Import-ModuleFromCache -ModuleName "Users"
                if ($moduleLoaded) {
                    Debug-ExcelData
                }
                else {
                    Write-LogMessage -Message "Failed to load Users module for debugging. Please check your internet connection." -Type Error
                }
                Read-Host "Press Enter to continue"
            }
            9 {
                # Refresh modules from GitHub
                Write-LogMessage -Message "Refreshing all modules from GitHub..." -Type Info
                $refreshCount = 0
                foreach ($moduleName in $GitHubConfig.ModuleFiles.Keys) {
                    $refreshed = Import-ModuleFromCache -ModuleName $moduleName -ForceRefresh
                    if ($refreshed) {
                        $refreshCount++
                    }
                }
                Write-LogMessage -Message "Successfully refreshed $refreshCount out of $($GitHubConfig.ModuleFiles.Count) modules" -Type Success
                Read-Host "Press Enter to continue"
            }
            10 {
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