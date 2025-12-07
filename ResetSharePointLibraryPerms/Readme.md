
# Reset Permissions (SharePoint files in a library)

## Overview
A PowerShell script that resets permissions on SharePoint files in a document library to inherit from the library.

## Prerequisites
- PowerShell 5.1 or higher
- PnP PowerShell module
- Entra ID app created with appropriate permissions to modify SharePoint sites
- Certificate of the Entra ID app installed in the user's CurrentUser\My store

## Installation
```powershell
Install-Module -Name PnP.PowerShell -Force
```

## Usage
```powershell
.\ResetPerms.ps1 -SiteUrl "https://yourtenant.sharepoint.com/sites/YourSite" -ClientId "your-app-id" -TenantId "your-tenant-id" -Thumbprint "ABC123DEF456..." -LibraryName "Documents"
```

## Parameters
- `-ClientId` - Entra ID App (Client) ID (required)
- `-TenantId` - Entra ID Tenant ID (GUID or name.onmicrosoft.com) (required)
- `-Thumbprint` - Certificate thumbprint from CurrentUser\My store (required)
- `-LibraryName` - The name of the document library to process, such as "Documents" (required)
- **SiteUrl** (required): The SharePoint site URL

## Features
- Resets file permissions to inherit from the library 

## Notes
- Always test in a non-production environment first
- Backup permission settings before running
