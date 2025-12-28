# List Files With Unique Permissions

## Description
`ListFilesWithUniquePermissions.ps1` is a PowerShell script that connects to Microsoft 365 and audits SharePoint sites to find files in document libraries that have unique permissions (i.e., they do not inherit from the library). It supports both CSV input and direct array parameters and outputs per-site CSVs plus per-library TXT files showing permissions.

## Features
- Connects to SharePoint Online via PnP PowerShell
- Enumerates each site's document libraries
- Detects files with unique permissions (`HasUniqueRoleAssignments`)
- Retrieves principal + role names for file-level permissions
- Includes sharing links fallback to capture items made unique via sharing
- Outputs one CSV per site and TXT permission files per library (only when that library has unique-permission files)

## Prerequisites
- PowerShell 7.4 or higher
- PnP PowerShell module installed (https://pnp.github.io/powershell/articles/installation.html)
- Entra ID app created with appropriate permissions to access SharePoint sites (https://pnp.github.io/powershell/articles/registerapplication.html)
- Certificate of the Entra ID app installed in the user's CurrentUser\My store (https://pnp.github.io/powershell/articles/authentication.html#non-interactive-authentication-using-a-certificate-in-the-windows-certificate-store)
- Provide sites via CSV `SiteUrl` column or `-SiteUrls` parameter

## Entra ID app permissions required
- SharePoint > Application Permissions > **Sites.FullControl.All**
- Make sure to grant Admin Consent for the permission

## Usage
Using a CSV of sites:
```powershell
.\nListFilesWithUniquePermissions.ps1 -ClientId "your-app-id" -TenantId "your-tenant-id" -Thumbprint "ABC123DEF456..." -SitesCsvPath "C:\Exports\sites.csv" -exportFolder "C:\Exports"
```

Using direct site URLs:
```powershell
.
ListFilesWithUniquePermissions.ps1 -ClientId "your-app-id" -TenantId "your-tenant-id" -Thumbprint "ABC123DEF456..." -SiteUrls @(
  "https://contoso.sharepoint.com/sites/ProjectA",
  "https://contoso.sharepoint.com/sites/ProjectB"
) -exportFolder "C:\Exports"
```

## Parameters
- `-ClientId` - Entra ID App (Client) ID (required)
- `-TenantId` - Entra ID Tenant ID (GUID or name.onmicrosoft.com) (required)
- `-Thumbprint` - Certificate thumbprint from CurrentUser\My store (required)
- `-SitesCsvPath` - Path to CSV file containing list of sites with `SiteUrl` column (required in CSV parameter set)
- `-SiteUrls` - Array of absolute HTTPS site URLs (required in DirectUrls parameter set)
- `-exportFolder` - Folder to write CSV and TXT outputs (required)

## Output
Per-site CSV with columns:
- `SiteUrl`
- `Library`
- `FileName`
- `FileUrl`
- `FilePermissions` (principal + roles; falls back to sharing links when role assignments are empty)
- `HasUniquePermissions`

Per-library TXT (only when unique-permission files exist in that library):
- Header lines for Site and Library
- `Permissions:` followed by entries like `PrincipalName (Role, Role)`; if the library inherits, effective site-level permissions are shown

## Support
For issues or questions, refer to Microsoft 365 documentation or PnP PowerShell documentation.

## Notes
- The `SiteUrl` column is required when using the CSV input
- URLs should be absolute HTTPS URLs
- Trailing slashes are optional
- Blank rows and invalid URLs are skipped with warnings
- Duplicate URLs are automatically removed
- TXT files are only created for libraries that contain files with unique permissions
