Here's your script — **`Get-SPOSearchInsights.ps1`** — ready to download. Here's a quick summary of what it does:

---

**What it does**
- Prompts you interactively (or accepts parameters) for:
  - Tenant Admin URL
  - All sites **or** a list of specific site URLs
  - **Last 28 days** or **Last 12 months**
- Authenticates via **PnP.PowerShell** (interactive browser login or app-only with a Client ID)
- Pulls search insights from **two sources** (for maximum coverage):
  1. **Microsoft Graph beta** – `getSiteSearchQueries`, `getSiteNoResultQueries`, `getSiteSearchAbandonedQueries`
  2. **SharePoint REST analytics API** per site – `GetTopQueries`, `GetNoResultQueries`, `GetAbandonedQueries`
  3. Falls back to the **Graph site usage report** if the search APIs return nothing (e.g. no analytics enabled)

**CSV columns exported**

| Column | Description |
|---|---|
| `ExportDate` | When the script ran |
| `ReportingPeriod` | `Last_28_Days` or `Last_12_Months` |
| `SiteTitle` / `SiteUrl` | Site identity |
| `InsightType` | TopQueries, NoResultQueries, AbandonedQueries, SiteUsageSummary |
| `Date` | Metric date (daily/monthly) |
| `QueryText` | The search query string |
| `QueryCount` | Times the query was run |
| `AbandonedCount` | Queries started but not completed |
| `NoResultsCount` | Queries with zero results |
| `ClickCount` | Result clicks |

**Requirements**
```powershell
Install-Module PnP.PowerShell -Scope CurrentUser
```
Your account needs **SharePoint Administrator** (or Global Admin) and the Graph `Reports.Read.All` permission.

**Usage examples**
```powershell
# All sites, interactive prompts
.\Get-SPOSearchInsights.ps1

# Specific sites, pass admin URL directly
.\Get-SPOSearchInsights.ps1 -TenantAdminUrl "https://contoso-admin.sharepoint.com" `
    -SiteUrls "https://contoso.sharepoint.com/sites/HR","https://contoso.sharepoint.com/sites/IT" `
    -OutputFolder "C:\Reports"
```