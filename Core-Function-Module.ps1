# === CoreFunctions.ps1 ===
# Core utility functions used throughout the Microsoft 365 Tenant Setup

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