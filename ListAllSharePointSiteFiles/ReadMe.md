
# List AllS harePoint Site Files

## Description
`ListAllSharePointSiteFiles.ps1` is a PowerShell script that connects to Microsoft 365 and retrieves a complete list of all files across provided SharePoint sites. It helps inventory and audit file storage.

## Features
- Connects to SharePoint Online via PnP PowerShell
- Enumerates all site's document libraries
- Retrieves files from document libraries
- Outputs results to CSV export per site

## Prerequisites
- PowerShell 5.1 or higher
- PnP PowerShell module installed
- Entra ID app created with appropriate permissions to access SharePoint sites
- Certificate of the Entra ID app installed in the user's CurrentUser\My store
- NOTE: edit the #sites array in the script to list all SharePoint sites you want to target. 

## Installation
```powershell
Install-Module PnP.PowerShell -Force
```

## Usage
```powershell
.\ListAllSharePointSiteFiles.ps1 -ClientId "your-app-id" -TenantId "your-tenant-id" -Thumbprint "ABC123DEF456..." -exportFolder "C:\Exports"
```

## Parameters
- `-ClientId` - Entra ID App (Client) ID (required)
- `-TenantId` - Entra ID Tenant ID (GUID or name.onmicrosoft.com) (required)
- `-Thumbprint` - Certificate thumbprint from CurrentUser\My store
- `-exportFolder` - FOLDER of the CSV to write to, such as "C:\Exports" (required)

## Output
Returns file details including:
- File name
- File path
- Size
- Last modified date
- Modified by

## Support
For issues or questions, refer to Microsoft 365 documentation or PnP PowerShell documentation.
