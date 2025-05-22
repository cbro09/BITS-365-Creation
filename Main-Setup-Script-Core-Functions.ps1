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

# === Import Module Functions ===
# Import all the individual module scripts
. "$PSScriptRoot\Modules\Groups.ps1"
. "$PSScriptRoot\Modules\ConditionalAccess.ps1"
. "$PSScriptRoot\Modules\SharePoint.ps1"
. "$PSScriptRoot\Modules\Intune.ps1"
. "$PSScriptRoot\Modules\Users.ps1"
. "$PSScriptRoot\Modules\Documentation.ps1"

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