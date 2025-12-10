# Created by Daniel Glenn October 28th, 2025; revised December 8th, 2025
# Repository: https://github.com/danielglenn/PowerShell-for-Microsoft-365

# Prerequisites:
# 1. Install the PnP.PowerShell module if not already installed:
#    Install-Module -Name PnP.PowerShell -Scope CurrentUser
# 2. Create an Entra ID App Registration with certificate-based authentication and appropriate SharePoint permissions.
#    Grant the app registration the necessary SharePoint permissions (e.g., Sites.Read.All) via the Entra ID portal.
# 3. Ensure you have the Entra ID App's certificate with a private key installed in the CurrentUser\My store on your computer and note its thumbprint.

# Usage:
# To call this script, this is the command:
#  .\GetSharePointSiteUsers.ps1 -SiteUrl "https://yourtenant.sharepoint.com/sites/YourSite" -ClientId "your-app-id" -TenantId "your-tenant-id" -Thumbprint "ABC123DEF456..."

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SiteUrl,
    
    [Parameter(Mandatory = $true)]
    [string]$ClientId,  # Entra ID App (Client) ID
    
    [Parameter(Mandatory = $true)]
    [string]$TenantId,  # Entra ID Tenant ID (GUID or name.onmicrosoft.com)
    
    [Parameter(Mandatory = $true)]
    [string]$Thumbprint  # Certificate thumbprint from CurrentUser\My store
)

# Import the PnP.PowerShell module
Import-Module PnP.PowerShell -ErrorAction Stop

# Validate thumbprint exists in CurrentUser\My store
Write-Host "Validating certificate thumbprint in CurrentUser\My store..." -ForegroundColor Cyan
$cert = Get-ChildItem -Path Cert:\CurrentUser\My -ErrorAction SilentlyContinue | Where-Object { $_.Thumbprint -eq $Thumbprint }
if (-not $cert) {
    Write-Error "Certificate with thumbprint '$Thumbprint' not found in CurrentUser\My store. Import it first (e.g., via Import-PfxCertificate)."
    return
}
Write-Host "Certificate validated: $($cert.Subject)" -ForegroundColor Green

# Define excluded system groups (case-insensitive)
$excludedGroups = @("Everyone", "All Users", "Everyone except external users")

# Function to get permissions for the site web (site-level only, excluding specified groups)
function Get-SitePermissions {
    param([Microsoft.SharePoint.Client.Web]$Web)

    $PermissionCollection = @()

    # Get site collection admins (they have Full Control)
    Write-Host "Fetching site collection administrators..." -ForegroundColor Yellow
    $SiteAdmins = Get-PnPSiteCollectionAdmin
    foreach ($Admin in $SiteAdmins) {
        $Permissions = New-Object PSObject
        $Permissions | Add-Member NoteProperty Title($Admin.Title)
        $Permissions | Add-Member NoteProperty Email($Admin.Email)
        $Permissions | Add-Member NoteProperty Permissions("Full Control")
        $Permissions | Add-Member NoteProperty GrantedThrough("Site Collection Administrator")
        $Permissions | Add-Member NoteProperty SecurityGroups("")
        $PermissionCollection += $Permissions
    }

    # Get role assignments for the web
    Get-PnPProperty -ClientObject $Web -Property RoleAssignments
    foreach ($RoleAssignment in $Web.RoleAssignments) {
        Get-PnPProperty -ClientObject $RoleAssignment -Property RoleDefinitionBindings, Member
        $PermissionLevels = ($RoleAssignment.RoleDefinitionBindings | Where-Object { $_.Name -ne "Limited Access" } | Select-Object -ExpandProperty Name) -join ", "
        
        if ($PermissionLevels.Length -eq 0) { continue }

        $PermissionType = $RoleAssignment.Member.PrincipalType
        $loginName = $RoleAssignment.Member.LoginName

        if ($PermissionType -eq "SharePointGroup") {
            # Check if this is an excluded system group
            $groupTitle = $RoleAssignment.Member.Title.ToLower()
            if ($excludedGroups | ForEach-Object { $_.ToLower() } | Where-Object { $_ -eq $groupTitle }) {
                Write-Host "Skipping excluded system group: $($RoleAssignment.Member.Title)" -ForegroundColor Gray
                continue
            }
            
            # Always add the group as a separate entry
            $GroupEntry = New-Object PSObject
            $GroupEntry | Add-Member NoteProperty Title($RoleAssignment.Member.Title)
            $GroupEntry | Add-Member NoteProperty Email("")
            $GroupEntry | Add-Member NoteProperty Permissions($PermissionLevels)
            $GroupEntry | Add-Member NoteProperty GrantedThrough("Security Group")
            $GroupEntry | Add-Member NoteProperty SecurityGroups("N/A")
            $PermissionCollection += $GroupEntry
            
            # Expand group to users
            $GroupMembers = Get-PnPGroupMember -Group $RoleAssignment.Member.Title -ErrorAction SilentlyContinue
            if ($GroupMembers.Count -gt 0) {
                foreach ($Member in $GroupMembers) {
                    $Permissions = New-Object PSObject
                    $Permissions | Add-Member NoteProperty Title($Member.Title)
                    $Permissions | Add-Member NoteProperty Email($Member.Email)
                    $Permissions | Add-Member NoteProperty Permissions($PermissionLevels)
                    $Permissions | Add-Member NoteProperty GrantedThrough("Group: $($RoleAssignment.Member.Title)")
                    $Permissions | Add-Member NoteProperty SecurityGroups($RoleAssignment.Member.Title)
                    $PermissionCollection += $Permissions
                }
            }
        } else {
            # Direct user or potential AAD group (PrincipalType = "User" or "SecurityGroup")
            Write-Host "Processing potential AAD/direct entry with LoginName: $loginName" -ForegroundColor Gray
            
            $isAADGroup = $false
            $group = $null
            $groupId = $null
            $userEmail = $RoleAssignment.Member.Email

            if ($loginName -match '^c:0o\.c\|federateddirectoryclaimprovider\|(.+)$') {
                # Parse AAD/M365 group GUID from claim format
                $groupId = $matches[1]
                try {
                    # Use PnP AzureAD cmdlets for AAD group fetch/expansion
                    $group = Get-PnPAzureADGroup -Identity $groupId -ErrorAction Stop
                    $isAADGroup = $true
                    Write-Host "Detected AAD group via GUID: $($group.DisplayName) ($groupId)" -ForegroundColor Green
                } catch {
                    Write-Warning "Failed to fetch AAD group with ID '$groupId': $($_.Exception.Message). Treating as direct user."
                }
            } elseif ($loginName -match '^i:0#\.f\|membership\|(.+)$') {
                # Parse direct user email from claim format
                $userEmail = $matches[1]
                Write-Host "Detected direct user via email: $userEmail" -ForegroundColor Gray
            }
            
            if ($isAADGroup -and $group) {
                # It's an AAD group (check exclusion)
                $groupTitle = $group.DisplayName.ToLower()
                if ($excludedGroups | ForEach-Object { $_.ToLower() } | Where-Object { $_ -eq $groupTitle }) {
                    Write-Host "Skipping excluded AAD group: $($group.DisplayName)" -ForegroundColor Gray
                    continue
                }

                # Add the group as a separate entry
                $GroupEntry = New-Object PSObject
                $GroupEntry | Add-Member NoteProperty Title($group.DisplayName)
                $GroupEntry | Add-Member NoteProperty Email($group.Mail)
                $GroupEntry | Add-Member NoteProperty Permissions($PermissionLevels)
                $GroupEntry | Add-Member NoteProperty GrantedThrough("AAD Security Group")
                $GroupEntry | Add-Member NoteProperty SecurityGroups("N/A")
                $PermissionCollection += $GroupEntry

                # Expand AAD group members using PnP
                $GroupMembers = Get-PnPAzureADGroupMember -Identity $group.Id -ErrorAction SilentlyContinue
                if ($GroupMembers.Count -gt 0) {
                    Write-Host "Expanding $($GroupMembers.Count) members from AAD group '$($group.DisplayName)'" -ForegroundColor Yellow
                    foreach ($Member in $GroupMembers) {
                        $memberUser = Get-PnPAzureADUser -Identity $Member.Id -ErrorAction SilentlyContinue
                        if ($memberUser) {
                            $Permissions = New-Object PSObject
                            $Permissions | Add-Member NoteProperty Title($memberUser.DisplayName)
                            $Permissions | Add-Member NoteProperty Email($memberUser.Mail)
                            $Permissions | Add-Member NoteProperty Permissions($PermissionLevels)
                            $Permissions | Add-Member NoteProperty GrantedThrough("Group: $($group.DisplayName)")
                            $Permissions | Add-Member NoteProperty SecurityGroups($group.DisplayName)
                            $PermissionCollection += $Permissions
                        }
                    }
                } else {
                    Write-Host "No members found for AAD group '$($group.DisplayName)'" -ForegroundColor Yellow
                }
            } else {
                # Fallback: Treat as direct user
                $Permissions = New-Object PSObject
                $Permissions | Add-Member NoteProperty Title($RoleAssignment.Member.Title)
                $Permissions | Add-Member NoteProperty Email($userEmail)
                $Permissions | Add-Member NoteProperty Permissions($PermissionLevels)
                $Permissions | Add-Member NoteProperty GrantedThrough("Direct Access")
                $Permissions | Add-Member NoteProperty SecurityGroups("")
                $PermissionCollection += $Permissions
            }
        }
    }

    # Deduplicate users/groups (combine permissions/groups if overlapping)
    $uniqueEntries = $PermissionCollection | Sort-Object Title | Group-Object Title | ForEach-Object {
        $entry = $_; $combinedPerms = ($entry.Group | ForEach-Object { $_.Permissions }) -join " | "
        $combinedGroups = ($entry.Group | Where-Object { $_.SecurityGroups -ne "N/A" } | ForEach-Object { $_.SecurityGroups }) -join " | "
        if ($combinedGroups -eq "") { $combinedGroups = "" } else { $combinedGroups = $combinedGroups }
        [PSCustomObject]@{
            Title = $entry.Name
            Email = ($entry.Group | Select-Object -First 1).Email
            Permissions = $combinedPerms
            GrantedThrough = ($entry.Group | ForEach-Object { $_.GrantedThrough }) -join " | "
            SecurityGroups = $combinedGroups
        }
    }

    return $uniqueEntries
}

try {
    # Connect to SharePoint using Thumbprint (app-only)
    Write-Host "Connecting to SharePoint site: $SiteUrl using Thumbprint auth..." -ForegroundColor Green
    Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Thumbprint $Thumbprint -Tenant $TenantId -ErrorAction Stop
    Write-Host "Successfully connected to SharePoint." -ForegroundColor Green

    # Get the web
    $Web = Get-PnPWeb

    # Retrieve users and groups with permissions (filtered, PnP-only AAD expansion)
    Write-Host "Fetching users, groups (SharePoint + AAD/M365 via PnP), and their permissions (excluding system groups)..." -ForegroundColor Yellow
    $entriesWithPerms = Get-SitePermissions -Web $Web

    if ($entriesWithPerms.Count -eq 0) {
        Write-Host "No users or groups found with access to the site (after exclusions)." -ForegroundColor Red
    } else {
        Write-Host "Found $($entriesWithPerms.Count) unique entries (users + groups) with access. Listing details below:" -ForegroundColor Green
        
        # Display in a formatted table (Title, Email, Permissions, SecurityGroups)
        $entriesWithPerms | Select-Object Title, Email, Permissions, SecurityGroups | Format-Table -AutoSize

        # Optional: Export to CSV
        $csvPath = "SharePointSiteEntriesWithGroupsAAD_PnPOnly_$((Get-Date).ToString('yyyyMMdd_HHmmss')).csv"
        $entriesWithPerms | Select-Object Title, Email, Permissions, SecurityGroups, GrantedThrough | Export-Csv -Path $csvPath -NoTypeInformation
        Write-Host "Exported full details (including groups and how granted) to: $csvPath" -ForegroundColor Green
    }
}
catch {
    $errorMsg = $_.Exception.Message
    if ($errorMsg -like "*thumbprint*" -or $errorMsg -like "*certificate*" -or $errorMsg -like "*authentication*") {
        Write-Error "Authentication failed. Check Thumbprint, ClientId, TenantId, and app permissions. Error: $errorMsg"
    } else {
        Write-Error "An error occurred: $errorMsg"
    }
}
finally {
    Disconnect-PnPOnline
    Write-Host "Disconnected from SharePoint." -ForegroundColor Cyan
}
