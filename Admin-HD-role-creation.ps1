# ===================================================================
# Azure AD Custom Role Creation Script
# Combines permissions from multiple built-in roles for helpdesk operations
# ===================================================================

# Prerequisites check and module installation
Write-Host "Checking prerequisites..." -ForegroundColor Cyan

# Check if running as Administrator (recommended for module installation)
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
if (-not $isAdmin) {
    Write-Warning "Not running as Administrator. Module installation may require elevation."
}

# Required modules
$requiredModules = @(
    "Microsoft.Graph.Authentication",
    "Microsoft.Graph.Identity.DirectoryManagement"
)

Write-Host "Checking required Microsoft Graph modules..." -ForegroundColor Yellow

foreach ($module in $requiredModules) {
    $installedModule = Get-Module -ListAvailable -Name $module
    
    if (-not $installedModule) {
        Write-Host "Module '$module' not found. Installing..." -ForegroundColor Red
        
        try {
            # Try to install for current user first (doesn't require admin)
            Install-Module -Name $module -Scope CurrentUser -Force -AllowClobber -Repository PSGallery
            Write-Host "Successfully installed $module for current user" -ForegroundColor Green
        }
        catch {
            Write-Host "Failed to install for current user. Trying AllUsers scope (requires admin)..." -ForegroundColor Yellow
            try {
                Install-Module -Name $module -Scope AllUsers -Force -AllowClobber -Repository PSGallery
                Write-Host "Successfully installed $module for all users" -ForegroundColor Green
            }
            catch {
                Write-Error "Failed to install module '$module': $($_.Exception.Message)"
                Write-Host "Manual installation required. Run: Install-Module -Name $module -Scope CurrentUser" -ForegroundColor Red
                exit 1
            }
        }
    }
    else {
        $latestVersion = $installedModule | Sort-Object Version -Descending | Select-Object -First 1
        Write-Host "Module '$module' found (Version: $($latestVersion.Version))" -ForegroundColor Green
        
        # Check if update is available
        try {
            $onlineVersion = Find-Module -Name $module -Repository PSGallery | Select-Object -ExpandProperty Version
            if ([version]$onlineVersion -gt [version]$latestVersion.Version) {
                Write-Host "Newer version available ($onlineVersion). Consider updating with: Update-Module -Name $module" -ForegroundColor Cyan
            }
        }
        catch {
            # Ignore if we can't check for updates
        }
    }
}

# Import required modules
Write-Host "Importing Microsoft Graph modules..." -ForegroundColor Yellow
foreach ($module in $requiredModules) {
    try {
        Import-Module -Name $module -Force
        Write-Host "Imported $module" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to import module '$module': $($_.Exception.Message)"
        exit 1
    }
}

# Connect to Microsoft Graph with required permissions
Write-Host "`nConnecting to Microsoft Graph..." -ForegroundColor Cyan

# Disconnect any existing connections to ensure fresh authentication
try {
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
}
catch {
    # Ignore if no existing connection
}

Write-Host "You will be prompted to sign in and consent to the required permissions." -ForegroundColor Gray

$requiredScopes = @(
    "RoleManagement.ReadWrite.Directory"
)

try {
    # Force fresh authentication
    Connect-MgGraph -Scopes $requiredScopes -NoWelcome
    
    # Verify connection and permissions
    $context = Get-MgContext
    if (-not $context) {
        Write-Error "Failed to connect to Microsoft Graph. Please check your credentials and permissions."
        exit 1
    }
    
    Write-Host "Connected successfully!" -ForegroundColor Green
    Write-Host "   Account: $($context.Account)" -ForegroundColor Gray
    Write-Host "   Tenant: $($context.TenantId)" -ForegroundColor Gray
    Write-Host "   Environment: $($context.Environment)" -ForegroundColor Gray
    
    # Verify we have the required permissions
    $currentScopes = $context.Scopes
    $missingScopes = $requiredScopes | Where-Object { $_ -notin $currentScopes }
    
    if ($missingScopes) {
        Write-Warning "Missing required permissions: $($missingScopes -join ', ')"
        Write-Host "You may need to disconnect and reconnect with admin consent." -ForegroundColor Yellow
        
        $response = Read-Host "Continue anyway? (y/N)"
        if ($response -ne 'y' -and $response -ne 'Y') {
            Write-Host "Exiting script. Please ensure you have the required permissions." -ForegroundColor Red
            Disconnect-MgGraph | Out-Null
            exit 1
        }
    }
    else {
        Write-Host "All required permissions granted" -ForegroundColor Green
    }
}
catch {
    Write-Error "Failed to connect to Microsoft Graph: $($_.Exception.Message)"
    Write-Host "`nTroubleshooting tips:" -ForegroundColor Yellow
    Write-Host "1. Ensure you have admin permissions in your Azure AD tenant" -ForegroundColor Gray
    Write-Host "2. Check if your account has 'Privileged Role Administrator' or 'Global Administrator' role" -ForegroundColor Gray
    Write-Host "3. Try running: Connect-MgGraph -Scopes 'RoleManagement.ReadWrite.Directory' -UseDeviceAuthentication" -ForegroundColor Gray
    exit 1
}

try {
    Write-Host "Creating custom role: 'Advanced Helpdesk Administrator v2'..." -ForegroundColor Yellow
    
    # First, check if role already exists
    Write-Host "Checking for existing roles with this name..." -ForegroundColor Gray
    $existingRole = Get-MgRoleManagementDirectoryRoleDefinition -Filter "displayName eq 'Advanced Helpdesk Administrator v2'"
    if ($existingRole) {
        Write-Host "Role already exists! ID: $($existingRole.Id)" -ForegroundColor Red
        $response = Read-Host "Delete existing role and recreate? (y/N)"
        if ($response -eq 'y' -or $response -eq 'Y') {
            Remove-MgRoleManagementDirectoryRoleDefinition -UnifiedRoleDefinitionId $existingRole.Id
            Write-Host "Existing role deleted." -ForegroundColor Green
        } else {
            Write-Host "Exiting. Use a different role name or delete the existing role." -ForegroundColor Yellow
            exit 0
        }
    }
    
    # Create role definition using proper JSON structure
    $roleDefinitionJson = @"
{
    "displayName": "Advanced Helpdesk Administrator v2",
    "description": "Custom role combining Authentication Administrator, User Administrator, Application Administrator, SharePoint Administrator, Office Apps Administrator, Teams Communications Administrator, and Global Reader permissions for comprehensive helpdesk operations",
    "isEnabled": true,
    "rolePermissions": [
        {
            "allowedResourceActions": [
                "microsoft.directory/users/authenticationMethods/create",
                "microsoft.directory/users/authenticationMethods/delete",
                "microsoft.directory/users/authenticationMethods/standard/read",
                "microsoft.directory/users/authenticationMethods/update",
                "microsoft.directory/users/strongAuthentication/update",
                "microsoft.directory/bitlockerKeys/key/read",
                "microsoft.directory/users/create",
                "microsoft.directory/users/delete",
                "microsoft.directory/users/disable",
                "microsoft.directory/users/enable",
                "microsoft.directory/users/invalidateAllRefreshTokens",
                "microsoft.directory/users/restore",
                "microsoft.directory/users/update",
                "microsoft.directory/users/assignLicense",
                "microsoft.directory/users/reprocessLicenseAssignment",
                "microsoft.directory/users/password/update",
                "microsoft.directory/users/userPrincipalName/update",
                "microsoft.directory/users/manager/update",
                "microsoft.directory/groups/create",
                "microsoft.directory/groups/delete",
                "microsoft.directory/groups/restore",
                "microsoft.directory/groups/members/update",
                "microsoft.directory/groups/owners/update",
                "microsoft.directory/groups/basic/update",
                "microsoft.directory/groups/assignLicense",
                "microsoft.directory/groups/reprocessLicenseAssignment",
                "microsoft.directory/applications/create",
                "microsoft.directory/applications/delete",
                "microsoft.directory/applications/basic/update",
                "microsoft.directory/applications/credentials/update",
                "microsoft.directory/applications/owners/update",
                "microsoft.directory/servicePrincipals/create",
                "microsoft.directory/servicePrincipals/delete",
                "microsoft.directory/servicePrincipals/disable",
                "microsoft.directory/servicePrincipals/enable",
                "microsoft.directory/servicePrincipals/basic/update",
                "microsoft.directory/servicePrincipals/credentials/update",
                "microsoft.directory/devices/delete",
                "microsoft.directory/devices/disable",
                "microsoft.directory/devices/enable",
                "microsoft.directory/devices/basic/update",
                "microsoft.directory/applications/allProperties/read",
                "microsoft.directory/auditLogs/allProperties/read",
                "microsoft.directory/contacts/allProperties/read",
                "microsoft.directory/devices/allProperties/read",
                "microsoft.directory/directoryRoles/allProperties/read",
                "microsoft.directory/domains/allProperties/read",
                "microsoft.directory/groups/allProperties/read",
                "microsoft.directory/organization/allProperties/read",
                "microsoft.directory/policies/allProperties/read",
                "microsoft.directory/roleAssignments/allProperties/read",
                "microsoft.directory/roleDefinitions/allProperties/read",
                "microsoft.directory/servicePrincipals/allProperties/read",
                "microsoft.directory/subscribedSkus/allProperties/read",
                "microsoft.directory/users/allProperties/read",
                "microsoft.directory/deletedItems/delete",
                "microsoft.directory/deletedItems/restore",
                "microsoft.office365.supportTickets/allEntities/allTasks",
                "microsoft.azure.supportTickets/allEntities/allTasks",
                "microsoft.office365.messageCenter/messages/read",
                "microsoft.office365.serviceHealth/allEntities/allTasks",
                "microsoft.office365.webPortal/allEntities/standard/read"
            ]
        }
    ]
}
"@

    # Convert JSON string to hashtable
    $roleDefinition = $roleDefinitionJson | ConvertFrom-Json -AsHashtable
    
    # Create the custom role
    $newRole = New-MgRoleManagementDirectoryRoleDefinition -BodyParameter $roleDefinition
    
    if (-not $newRole -or -not $newRole.Id) {
        throw "Role creation failed - no role ID returned"
    }
    
    Write-Host "Custom role created successfully!" -ForegroundColor Green
    Write-Host "Role ID: $($newRole.Id)" -ForegroundColor Cyan
    Write-Host "Role Name: $($newRole.DisplayName)" -ForegroundColor Cyan
    Write-Host "Total Permissions: $($newRole.RolePermissions[0].AllowedResourceActions.Count)" -ForegroundColor Cyan
    
    # Display role summary
    Write-Host "`nRole Summary:" -ForegroundColor White
    Write-Host "* Authentication management (MFA, password reset, auth methods)" -ForegroundColor Gray
    Write-Host "* User lifecycle management (create, update, delete, licensing)" -ForegroundColor Gray  
    Write-Host "* Group management (create, update, members, licensing)" -ForegroundColor Gray
    Write-Host "* Application management (enterprise apps, registrations)" -ForegroundColor Gray
    Write-Host "* Device management (enable, disable, update)" -ForegroundColor Gray
    Write-Host "* Comprehensive read access across all directory objects" -ForegroundColor Gray
    Write-Host "* Support ticket creation and management" -ForegroundColor Gray
    
    Write-Host "`nImportant Notes:" -ForegroundColor Yellow
    Write-Host "1. This role provides extensive permissions - review before assignment" -ForegroundColor Red
    Write-Host "2. For Intune device management, users also need the 'Help Desk Operator' role in Intune" -ForegroundColor Yellow  
    Write-Host "3. Consider implementing Conditional Access policies for this role" -ForegroundColor Yellow
    Write-Host "4. Use Privileged Identity Management (PIM) for just-in-time access" -ForegroundColor Yellow
    
    Write-Host "`nNext Steps:" -ForegroundColor White
    Write-Host "1. Review role permissions in Azure AD Portal > Roles and administrators" -ForegroundColor Gray
    Write-Host "2. Test with a pilot user before broad deployment" -ForegroundColor Gray
    Write-Host "3. Assign both this role AND Intune 'Help Desk Operator' role to users" -ForegroundColor Gray
    Write-Host "4. Configure appropriate scope limitations if needed" -ForegroundColor Gray
    
    # Optional: Display role assignment script
    Write-Host "`nTo assign this role to a user, use:" -ForegroundColor Cyan
    Write-Host @"
`$userId = "user@yourdomain.com"
`$roleId = "$($newRole.Id)"
`$params = @{
    "@odata.type" = "#microsoft.graph.unifiedRoleAssignment"
    principalId = (Get-MgUser -Filter "userPrincipalName eq '`$userId'").Id
    roleDefinitionId = `$roleId
    directoryScopeId = "/"
}
New-MgRoleManagementDirectoryRoleAssignment -BodyParameter `$params
"@ -ForegroundColor Gray

} catch {
    Write-Error "Failed to create custom role: $($_.Exception.Message)"
    Write-Host "`nDetailed error information:" -ForegroundColor Yellow
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Category: $($_.CategoryInfo.Category)" -ForegroundColor Red
    Write-Host "Activity: $($_.CategoryInfo.Activity)" -ForegroundColor Red
    
    Write-Host "`nCommon issues:" -ForegroundColor Yellow
    Write-Host "* Insufficient permissions (need Privileged Role Administrator or Global Administrator)" -ForegroundColor Red
    Write-Host "* Role name already exists (we should have caught this)" -ForegroundColor Red
    Write-Host "* Invalid permission specified (some permissions may not be available in your tenant)" -ForegroundColor Red
    Write-Host "* API structure error (JSON formatting issue)" -ForegroundColor Red
    
    Write-Host "`nTroubleshooting steps:" -ForegroundColor Cyan
    Write-Host "1. Verify you have Global Administrator or Privileged Role Administrator role" -ForegroundColor Gray
    Write-Host "2. Check if any similar role names exist in Azure AD Portal > Roles and administrators" -ForegroundColor Gray
    Write-Host "3. Try running with fewer permissions to identify problematic ones" -ForegroundColor Gray
    
    exit 1
}

Write-Host "`nScript completed!" -ForegroundColor Green