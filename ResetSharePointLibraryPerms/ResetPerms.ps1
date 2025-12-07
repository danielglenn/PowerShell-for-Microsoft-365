# Created by Daniel Glenn September 20, 2025, revised December 6th, 2025
# Repository: https://github.com/danielglenn/PowerShell-for-Microsoft-365
# Used to reset permissions of files in a document library to inherit permissions from the library where they are located. This is done for many reasons, 
# including so content can be manually archived and not worry that users still have access to documents via individual permissions.

# Prerequisites:
# 1. Install the PnP.PowerShell module if not already installed:
#    Install-Module -Name PnP.PowerShell -Scope CurrentUser
# 2. Create an Entra ID App Registration with certificate-based authentication and appropriate SharePoint permissions.
#    Grant the app registration the necessary SharePoint permissions (e.g., Sites.ReadWrite.All) via the Entra ID portal.
# 3. Ensure you have a certificate with a private key installed in the CurrentUser\My store on your computer and note its thumbprint.

# Usage:
# To call this script, use the following syntax, replacing the parameters with your own values:
#  .\ResetPerms.ps1 -SiteUrl "https://yourtenant.sharepoint.com/sites/YourSite" -ClientId "your-app-id" -TenantId "your-tenant-id" -Thumbprint "ABC123DEF456..." -LibraryName "Documents"


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

    [Parameter(Mandatory = $true)]
    [string]$LibraryName  # The name of the document library to process, such as "Documents"
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

#If needed, connect to SPO:  
connect-pnponline -url $SiteUrl -clientid $ClientId -tenant $TenantId -Thumbprint $Thumbprint

$items = Get-PnPListItem -list $libraryName -PageSize 1000
foreach ($item in $items) {
$hasUniquePermissions = Get-PnPProperty -ClientObject $item -Property HasUniqueRoleAssignments
if ($hasUniquePermissions) {
$item.ResetRoleInheritance()
$item.Context.ExecuteQuery()
Write-Host "Reset Permissions for item ID $($item.Id)"
} else {
Write-Host "Item ID $($item.Id) already inherits permissions."
}
}