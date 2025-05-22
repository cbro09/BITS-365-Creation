#requires -Version 5.1
<#
.SYNOPSIS
    Main Menu for Microsoft 365 Tenant Setup Utility
.DESCRIPTION
    Provides a unified menu interface for Microsoft 365 tenant setup operations
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

# === Import Module Functions ===
# Import all the individual module scripts
. "$PSScriptRoot\Modules\CoreFunctions.ps1"
. "$PSScriptRoot\Modules\GraphConnection.ps1"
. "$PSScriptRoot\Modules\Groups.ps1"
. "$PSScriptRoot\Modules\ConditionalAccess.ps1"
. "$PSScriptRoot\Modules\SharePoint.ps1"
. "$PSScriptRoot\Modules\Intune.ps1"
. "$PSScriptRoot\Modules\Users.ps1"
. "$PSScriptRoot\Modules\Documentation.ps1"

# === Main Menu Functions ===
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