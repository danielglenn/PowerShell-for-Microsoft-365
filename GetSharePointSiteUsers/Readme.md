
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
- PnP.PowerShell module installed (https://pnp.github.io/powershell/articles/installation.html)
- Entra ID app registration with SharePoint API permissions (https://pnp.github.io/powershell/articles/registerapplication.html)
- Client certificate installed in `CurrentUser\My` store (https://pnp.github.io/powershell/articles/authentication.html#non-interactive-authentication-using-a-certificate-in-the-windows-certificate-store)


## Entra ID app permissions required
**SharePoint** (application permission)
- Sites.Read.All — read site/web/role assignments and site collection admins (If you prefer to give full control instead of read-only) Sites.FullControl.All

**Microsoft Graph** (application permissions)
- Group.Read.All — read AAD/M365 groups and list group members
- User.Read.All — read user profile properties (displayName, mail, etc.)
- (Alternative: Directory.Read.All covers both groups and users but is broader — only use it if you need directory-wide read access.)

**Make sure to grant Admin Consent for the permissions.**

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
