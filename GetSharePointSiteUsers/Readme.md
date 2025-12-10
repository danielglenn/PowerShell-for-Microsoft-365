
# Get SharePoint Site Users

`GetSharePointSiteUsers.ps1` is a PowerShell script that retrieves and lists all users and security groups with access to a SharePoint site using certificate-based authentication (app-only).

## Features

- **App-only authentication** using certificate thumbprint
- **Expands security groups** (both SharePoint and Azure AD/M365)
- **Excludes system groups** (Everyone, All Users, etc.)
- **Deduplicates entries** with combined permissions
- **Site collection admins** identification
- **CSV export** with full permission details

## Prerequisites

- [PnP.PowerShell](https://pnp.github.io/powershell/) module installed
- Entra ID app registration with SharePoint API permissions
- Client certificate installed in `CurrentUser\My` store

## Usage

```powershell
.\GetSharePointSiteUsers.ps1 `
    -SiteUrl "https://yourtenant.sharepoint.com/sites/YourSite" `
    -ClientId "your-app-id" `
    -TenantId "your-tenant-id" `
    -Thumbprint "ABC123DEF456..."
```

## Parameters

| Parameter | Description |
|-----------|-------------|
| `SiteUrl` | SharePoint site URL |
| `ClientId` | Entra ID app (client) ID |
| `TenantId` | Entra ID tenant ID (GUID or name.onmicrosoft.com) |
| `Thumbprint` | Certificate thumbprint from CurrentUser\My store |

## Output

- Console table with Title, Email, Permissions, and SecurityGroups
- CSV file: `SharePointSiteEntriesWithGroupsAAD_PnPOnly_[timestamp].csv`
