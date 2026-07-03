#Requires -Version 7.4
#Requires -Modules @{ ModuleName = 'PnP.PowerShell'; ModuleVersion = '2.0.0' }
<#
.SYNOPSIS
    Exports SharePoint Online site usage insights to CSV.
.DESCRIPTION
    Connects to SharePoint Online and retrieves site usage report data for all
    SharePoint sites or a specific set of site URLs. Results are exported to a
    well-formatted CSV file.
    Requires:
        - PowerShell 7.4 or later
        - PnP.PowerShell module  (Install-Module PnP.PowerShell -Scope CurrentUser)
        - An account with SharePoint Administrator or Global Administrator role
        - Microsoft Graph permissions: Reports.Read.All, Sites.Read.All
    Note: SharePoint Online does not expose the site-level _api/search/analytics
    REST endpoints used by older SharePoint implementations, so this script
    uses the Microsoft Graph site usage report.

    Note: If the Microsoft 365 admin center setting "Reports display concealed
    user, group, and site names" is enabled (Settings > Org settings > Services >
    Reports), the site usage report returns anonymized values for Site URL and
    Owner Display Name instead of real names. When that setting is on, the
    SiteUrl column will contain masked identifiers and the -SiteUrls filter will
    not match real site URLs. Disable that setting to export real site names and
    URLs.
.PARAMETER TenantAdminUrl
    Your SharePoint Admin Centre URL.
    Example: https://contoso-admin.sharepoint.com
.PARAMETER SiteUrls
    Optional. One or more site URLs to query.
    If omitted, every non-OneDrive SharePoint site is queried.
    Example: @('https://contoso.sharepoint.com/sites/HR','https://contoso.sharepoint.com/sites/IT')
.PARAMETER AllSites
    Optional switch to run against all SharePoint sites (excluding OneDrive)
    without prompting for site scope.
.PARAMETER ReportPeriod
    Optional reporting period.
    Accepted values: Last28Days or Last12Months.
    If omitted, the script prompts for the reporting period.
.PARAMETER NoDedupe
    Optional switch to disable row de-duplication.
    Useful for very large exports when preserving every row and minimizing
    in-memory dedupe keys is preferred.
.PARAMETER OutputFolder
    Folder where the CSV file is written.  Defaults to the current directory.
.PARAMETER ClientId
    Azure AD App (client) ID used for certificate-based app-only auth.
.PARAMETER TenantId
    Microsoft Entra tenant ID (GUID) used for certificate-based app-only auth.
.PARAMETER Thumbprint
    Certificate thumbprint from the local certificate store used for app-only auth.
.EXAMPLE
    .\Get-SPOSearchInsights.ps1 `
        -TenantAdminUrl "https://contoso-admin.sharepoint.com" `
        -ClientId "11111111-2222-3333-4444-555555555555" `
        -TenantId "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee" `
        -Thumbprint "ABCDEF1234567890ABCDEF1234567890ABCDEF12" `
        -AllSites `
        -ReportPeriod Last28Days `
        -OutputFolder "C:\Reports"
.EXAMPLE
    .\Get-SPOSearchInsights.ps1 `
        -TenantAdminUrl "https://contoso-admin.sharepoint.com" `
        -ClientId "11111111-2222-3333-4444-555555555555" `
        -TenantId "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee" `
        -Thumbprint "ABCDEF1234567890ABCDEF1234567890ABCDEF12" `
        -ReportPeriod Last12Months
.EXAMPLE
    .\Get-SPOSearchInsights.ps1 `
        -TenantAdminUrl "https://contoso-admin.sharepoint.com" `
        -ClientId "11111111-2222-3333-4444-555555555555" `
        -TenantId "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee" `
        -Thumbprint "ABCDEF1234567890ABCDEF1234567890ABCDEF12" `
        -ReportPeriod Last28Days `
        -SiteUrls @(
            "https://contoso.sharepoint.com/sites/Marketing",
            "https://contoso.sharepoint.com/sites/HR",
            "https://contoso.sharepoint.com/sites/IT"
        ) `
        -OutputFolder "C:\Reports"
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$TenantAdminUrl,
    [Parameter()]
    [string[]]$SiteUrls,
    [Parameter()]
    [switch]$AllSites,
    [Parameter()]
    [ValidateSet('Last28Days', 'Last12Months')]
    [string]$ReportPeriod,
    [Parameter()]
    [switch]$NoDedupe,
    [Parameter()]
    [string]$OutputFolder = (Get-Location).Path,
    [Parameter(Mandatory = $true)]
    [string]$ClientId,
    [Parameter(Mandatory = $true)]
    [string]$TenantId,
    [Parameter(Mandatory = $true)]
    [string]$Thumbprint
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
#region ── Helpers ────────────────────────────────────────────────────────────
function Write-Header {
    param([string]$Text)
    Write-Host ""
    Write-Host ("=" * 60) -ForegroundColor Cyan
    Write-Host "  $Text" -ForegroundColor Cyan
    Write-Host ("=" * 60) -ForegroundColor Cyan
}
function Write-Step {
    param([string]$Text)
    Write-Host "  ► $Text" -ForegroundColor Yellow
}
function Write-OK {
    param([string]$Text)
    Write-Host "  ✔ $Text" -ForegroundColor Green
}
function Write-Warn {
    param([string]$Text)
    Write-Host "  ⚠ $Text" -ForegroundColor DarkYellow
}
function Write-Fail {
    param([string]$Text)
    Write-Host "  ✘ $Text" -ForegroundColor Red
}
function Get-JwtClaims {
    # Decodes the payload of a JWT (no signature validation) so we can inspect
    # the audience and application roles actually present on an access token.
    param([string]$Token)
    if ([string]::IsNullOrWhiteSpace($Token)) { return $null }
    $parts = $Token.Split('.')
    if ($parts.Count -lt 2) { return $null }
    $payload = $parts[1].Replace('-', '+').Replace('_', '/')
    switch ($payload.Length % 4) {
        2 { $payload += '==' }
        3 { $payload += '=' }
    }
    try {
        $json = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String($payload))
        return $json | ConvertFrom-Json
    } catch {
        return $null
    }
}
function Get-CsvValue {
    # Safely read a named column from a ConvertFrom-Csv row; returns '' when the
    # column is absent (StrictMode would otherwise throw on a missing property).
    param($Row, [string]$Name)
    $prop = $Row.PSObject.Properties[$Name]
    if ($prop -and $null -ne $prop.Value) { return [string]$prop.Value }
    return ''
}
#endregion
#region ── Module check ───────────────────────────────────────────────────────
Write-Header "SharePoint Online Site Usage Insights Exporter"
if (-not (Get-Module -ListAvailable -Name 'PnP.PowerShell' | Where-Object { $_.Version -ge [Version]'2.0' })) {
    Write-Fail "PnP.PowerShell v2+ is required."
    Write-Host "  Install it with:  Install-Module PnP.PowerShell -Scope CurrentUser" -ForegroundColor Gray
    exit 1
}
Import-Module PnP.PowerShell -MinimumVersion '2.0.0' -ErrorAction Stop
#endregion
#region ── Interactive prompts ────────────────────────────────────────────────
$TenantAdminUrl = $TenantAdminUrl.TrimEnd('/')

if ($AllSites -and $SiteUrls -and $SiteUrls.Count -gt 0) {
    Write-Fail "Use either -AllSites or -SiteUrls, not both."
    exit 1
}

# Site scope
if (-not $SiteUrls -and -not $AllSites) {
    Write-Host ""
    Write-Host "  Which sites do you want to query?" -ForegroundColor Yellow
    Write-Host "  [1] All SharePoint sites (excludes OneDrive)"
    Write-Host "  [2] Specific sites (you will be prompted)"
    do { $scopeChoice = Read-Host "  Choice (1 or 2)" } while ($scopeChoice -notin '1','2')
    if ($scopeChoice -eq '2') {
        Write-Host "  Enter site URLs one per line.  Leave blank and press Enter when done." -ForegroundColor Gray
        $collected = [System.Collections.Generic.List[string]]::new()
        while ($true) {
            $line = (Read-Host "  Site URL").Trim()
            if ([string]::IsNullOrEmpty($line)) { break }
            $collected.Add($line.TrimEnd('/'))
        }
        if ($collected.Count -eq 0) {
            Write-Fail "No sites entered.  Exiting."
            exit 1
        }
        $SiteUrls = $collected.ToArray()
    }
}
# Time period
$selectedReportPeriod = $ReportPeriod
if (-not $selectedReportPeriod) {
    Write-Host ""
    Write-Host "  Select the reporting period:" -ForegroundColor Yellow
    Write-Host "  [1] Last 28 days"
    Write-Host "  [2] Last 12 months"
    do { $periodChoice = Read-Host "  Choice (1 or 2)" } while ($periodChoice -notin '1','2')
    $selectedReportPeriod = if ($periodChoice -eq '1') { 'Last28Days' } else { 'Last12Months' }
}

$endDate   = (Get-Date).Date
if ($selectedReportPeriod -eq 'Last28Days') {
    $startDate       = $endDate.AddDays(-28)
    $periodLabel     = "Last_28_Days"
} else {
    $startDate       = $endDate.AddMonths(-12)
    $periodLabel     = "Last_12_Months"
}
$startDateStr = $startDate.ToString("yyyy-MM-dd")
$endDateStr   = $endDate.ToString("yyyy-MM-dd")
Write-Host ""
Write-OK "Period : $startDateStr → $endDateStr  ($periodLabel)"
#endregion
#region ── Connect ────────────────────────────────────────────────────────────
Write-Header "Connecting to SharePoint Online"
Write-Step "Authenticating with app-only certificate credentials …"

if ([string]::IsNullOrWhiteSpace($ClientId) -or
    [string]::IsNullOrWhiteSpace($TenantId) -or
    [string]::IsNullOrWhiteSpace($Thumbprint)) {
    Write-Fail "ClientId, TenantId, and Thumbprint are required for app-only auth."
    exit 1
}

$connectParams = @{
    Url         = $TenantAdminUrl
    ClientId    = $ClientId
    Tenant      = $TenantId
    Thumbprint  = $Thumbprint
}

try {
    Connect-PnPOnline @connectParams
    Write-OK "Connected to $TenantAdminUrl"
} catch {
    Write-Fail "Connection failed: $_"
    exit 1
}
#endregion
#region ── Data collection ────────────────────────────────────────────────────
Write-Header "Collecting Site Usage Insights"
$columnOrder = @(
    'ExportDate','ReportingPeriod','Source','SiteTitle','SiteUrl',
    'InsightType','Date','QueryText',
    'QueryCount','AbandonedCount','NoResultsCount','ClickCount'
)

if (-not (Test-Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
}

$timestamp       = Get-Date -Format 'yyyyMMdd_HHmmss'
$csvFile         = Join-Path $OutputFolder "SPO_SiteUsage_${periodLabel}_${timestamp}.csv"
$bufferFlushSize = 500
$rowBuffer       = [System.Collections.Generic.List[PSCustomObject]]::new()
$seenRowKeys     = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$runExportDate   = Get-Date -Format 'yyyy-MM-dd HH:mm'
$totalRows       = 0
$insightCounts   = @{}
$siteCounts      = @{}
$tokenHasReportsReadAll = $false
$usageRowsTotal = 0
$usageRowsWithSiteUrl = 0
$usageRowsMatchedScope = 0

function Clear-RowBuffer {
    if ($rowBuffer.Count -eq 0) { return }
    $append = Test-Path $csvFile
    $rowBuffer |
        Select-Object $columnOrder |
        Export-Csv -Path $csvFile -NoTypeInformation -Encoding UTF8 -Append:$append
    $rowBuffer.Clear()
}

function Add-ResultRows {
    param([object[]]$Rows)

    foreach ($row in $Rows) {
        if (-not $row) { continue }

        $source      = if ($row.PSObject.Properties['Source'] -and $row.Source) { [string]$row.Source } else { 'Unknown' }
        $siteUrlNorm = if ($row.SiteUrl) { $row.SiteUrl.TrimEnd('/') } else { '' }
        $insightType = if ($row.InsightType) { [string]$row.InsightType } else { '' }
        $dateValue   = if ($row.Date) { [string]$row.Date } else { '' }
        $queryText   = if ($row.QueryText) { [string]$row.QueryText } else { '' }
        $queryCount  = if ($row.QueryCount) { [string]$row.QueryCount } else { '' }
        $abandoned   = if ($row.AbandonedCount) { [string]$row.AbandonedCount } else { '' }
        $noResults   = if ($row.NoResultsCount) { [string]$row.NoResultsCount } else { '' }
        $clickCount  = if ($row.ClickCount) { [string]$row.ClickCount } else { '' }
        $dedupeKey   = "$source|$siteUrlNorm|$insightType|$dateValue|$queryText|$queryCount|$abandoned|$noResults|$clickCount"

        if (-not $NoDedupe -and -not $seenRowKeys.Add($dedupeKey)) { continue }

        $row.SiteUrl = $siteUrlNorm
        $row.Source = $source
        $row.ExportDate = $runExportDate
        $rowBuffer.Add($row)
        $script:totalRows++

        if (-not $insightCounts.ContainsKey($insightType)) {
            $insightCounts[$insightType] = 0
        }
        $insightCounts[$insightType]++

        $siteTitle = if ($row.SiteTitle) { [string]$row.SiteTitle } else { '(Unknown)' }
        $siteKey = if ($siteUrlNorm) { $siteUrlNorm } else { $siteTitle }
        if (-not $siteCounts.ContainsKey($siteKey)) {
            $siteCounts[$siteKey] = 0
        }
        $siteCounts[$siteKey]++

        if ($rowBuffer.Count -ge $bufferFlushSize) {
            Clear-RowBuffer
        }
    }
}
# Retrieve the tenant-wide SharePoint site usage report (page views and site
# activity per site) and filter to the requested site scope.
$usagePeriod = if ($selectedReportPeriod -eq 'Last28Days') { 'D30' } else { 'D180' }

# Inspect the Graph token so permission problems are obvious rather than opaque.
try {
    $graphToken = Get-PnPAccessToken -ResourceTypeName Graph
    $claims     = Get-JwtClaims -Token $graphToken
    if ($claims) {
        $roles = if ($claims.PSObject.Properties['roles'] -and $claims.roles) { ($claims.roles -join ', ') } else { '(none)' }
        $tokenHasReportsReadAll = $roles -match 'Reports\.Read\.All'
        Write-Step "Graph token audience: $($claims.aud)"
        Write-Step "Graph token app roles: $roles"
        if (-not $tokenHasReportsReadAll) {
            Write-Warn "The Graph token is missing the Reports.Read.All application role."
            Write-Warn "In Entra: App registration > API permissions > Microsoft Graph > Application permissions > add Reports.Read.All, then Grant admin consent."
        }
    } else {
        Write-Warn "Could not decode the Graph access token to verify permissions."
    }
} catch {
    Write-Warn "Could not acquire a Graph access token: $($_.Exception.Message)"
}

Write-Step "Retrieving SharePoint Site Usage report (period '$usagePeriod') …"
try {
    $usageUrl  = "v1.0/reports/getSharePointSiteUsageDetail(period='$usagePeriod')"
    $usageCsv  = Invoke-PnPGraphMethod -Url $usageUrl -Method Get -Raw
    $usageRows = $usageCsv | ConvertFrom-Csv
    $usageRowsTotal = @($usageRows).Count
    foreach ($row in $usageRows) {
        $siteUrl = (Get-CsvValue $row 'Site URL') -replace '/$',''
        if (-not $siteUrl) { continue }
        $usageRowsWithSiteUrl++
        $included = (-not $SiteUrls -or $SiteUrls.Count -eq 0) -or ($SiteUrls | Where-Object { $_.TrimEnd('/') -eq $siteUrl })
        if (-not $included) { continue }
        $usageRowsMatchedScope++
        $ownerName = Get-CsvValue $row 'Owner Display Name'
        $siteTitle = if ($siteUrl) { $siteUrl.TrimEnd('/').Split('/')[-1] } else { $ownerName }
        Add-ResultRows -Rows @([PSCustomObject]@{
            ExportDate      = $runExportDate
            ReportingPeriod = $periodLabel
            Source          = 'GRAPH_USAGE'
            SiteTitle       = $siteTitle
            SiteUrl         = $siteUrl
            InsightType     = 'SiteUsageSummary'
            Date            = (Get-CsvValue $row 'Report Refresh Date')
            QueryText       = ''
            QueryCount      = (Get-CsvValue $row 'Visited Page Count')
            AbandonedCount  = ''
            NoResultsCount  = ''
            ClickCount      = (Get-CsvValue $row 'Page View Count')
        })
    }
    Write-Step "Usage rows returned: $usageRowsTotal; rows with Site URL: $usageRowsWithSiteUrl; rows matching scope: $usageRowsMatchedScope"
    Write-OK "Site usage report: $totalRows row(s) collected."
} catch {
    Write-Fail "Site usage report failed: $_"
}
#endregion
#region ── Export ─────────────────────────────────────────────────────────────
Write-Header "Exporting Results"
if ($totalRows -eq 0) {
    Write-Warn "No data was collected.  The CSV will not be created."
    Write-Host ""
    Write-Host "  Possible reasons:" -ForegroundColor Gray
    if (-not $tokenHasReportsReadAll) {
        Write-Host "  • Your app token does not include Reports.Read.All Graph application permission." -ForegroundColor Gray
    }
    if ($usageRowsTotal -eq 0) {
        Write-Host "  • The usage report returned no rows for the selected period." -ForegroundColor Gray
    }
    if ($usageRowsTotal -gt 0 -and $usageRowsWithSiteUrl -eq 0) {
        Write-Host "  • Site URL values are concealed/anonymized in tenant report settings." -ForegroundColor Gray
    }
    if ($usageRowsWithSiteUrl -gt 0 -and $usageRowsMatchedScope -eq 0 -and $SiteUrls -and $SiteUrls.Count -gt 0) {
        Write-Host "  • None of the returned site URLs matched -SiteUrls (check concealment setting or URL normalization)." -ForegroundColor Gray
    }
    if ($usageRowsMatchedScope -gt 0 -and $totalRows -eq 0) {
        Write-Host "  • Rows were matched but filtered out (for example, de-duplication removed duplicates)." -ForegroundColor Gray
    }
} else {
    Clear-RowBuffer
    Write-OK "Exported $totalRows row(s) to:"
    Write-Host "  $csvFile" -ForegroundColor White
    # Quick summary table
    Write-Host ""
    Write-Host "  Summary by InsightType:" -ForegroundColor Cyan
    Write-Host ("  {0,-25} {1,8}" -f "Insight Type", "Rows") -ForegroundColor Cyan
    Write-Host ("  {0,-25} {1,8}" -f ("-" * 25), ("-" * 8)) -ForegroundColor Cyan
    $insightCounts.GetEnumerator() |
        Sort-Object Name |
        ForEach-Object {
            Write-Host ("  {0,-25} {1,8}" -f $_.Name, $_.Value)
        }
    Write-Host ""
    Write-Host "  Summary by Site:" -ForegroundColor Cyan
    Write-Host ("  {0,-40} {1,8}" -f "Site", "Rows") -ForegroundColor Cyan
    Write-Host ("  {0,-40} {1,8}" -f ("-" * 40), ("-" * 8)) -ForegroundColor Cyan
    $siteCounts.GetEnumerator() |
        Sort-Object Value -Descending |
        Select-Object -First 20 |
        ForEach-Object {
            $name = if ($_.Key.Length -gt 38) { $_.Key.Substring(0,37) + '…' } else { $_.Key }
            Write-Host ("  {0,-40} {1,8}" -f $name, $_.Value)
        }
}
Write-Host ""
Write-Host ("=" * 60) -ForegroundColor Cyan
Write-Host "  Done." -ForegroundColor Cyan
Write-Host ("=" * 60) -ForegroundColor Cyan
Write-Host ""
#endregion
#region ── Column reference ───────────────────────────────────────────────────
<#
  CSV Column Reference
  ─────────────────────────────────────────────────────────────
  ExportDate       Date/time the script was run
  ReportingPeriod  "Last_28_Days" or "Last_12_Months"
  Source           Data source (GRAPH_USAGE)
  SiteTitle        Display name of the site
  SiteUrl          Absolute URL of the site
  InsightType      SiteUsageSummary
  Date             Date the metric relates to (report refresh date)
  QueryText        Blank for summary rows
  QueryCount       Visited page count
  AbandonedCount   Blank for summary rows
  NoResultsCount   Blank for summary rows
  ClickCount       Page view count
#>
#endregion