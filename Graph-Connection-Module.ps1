# === GraphConnection.ps1 ===
# Microsoft Graph connection and tenant verification functions

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