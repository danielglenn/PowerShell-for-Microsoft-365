#Requires -Version 7.4
#Requires -Modules @{ ModuleName = 'PnP.PowerShell'; ModuleVersion = '2.0.0' }
<#
.SYNOPSIS
    Exports SharePoint Online Search Insights to CSV.
.DESCRIPTION
    Connects to SharePoint Online and retrieves Search Insights (top queries,
    no-result queries, abandoned queries, and query volume) for all SharePoint
    sites or a specific set of site URLs.  Results are exported to a
    well-formatted CSV file.
    Requires:
        - PowerShell 7.4 or later
        - PnP.PowerShell module  (Install-Module PnP.PowerShell -Scope CurrentUser)
        - An account with SharePoint Administrator or Global Administrator role
        - Microsoft Graph permissions: Reports.Read.All, Sites.Read.All
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
.PARAMETER HeartbeatIntervalSites
    Optional heartbeat interval for site processing status messages.
    A message is written every N sites (default: 50).
    Set to 0 to disable heartbeat messages.
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
    [ValidateRange(0, 100000)]
    [int]$HeartbeatIntervalSites = 50,
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
#endregion
#region ── Module check ───────────────────────────────────────────────────────
Write-Header "SharePoint Online Search Insights Exporter"
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
    $aggInterval     = "Daily"
} else {
    $startDate       = $endDate.AddMonths(-12)
    $periodLabel     = "Last_12_Months"
    $aggInterval     = "Monthly"
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
#region ── Discover sites ─────────────────────────────────────────────────────
Write-Header "Discovering Sites"
if ($SiteUrls -and $SiteUrls.Count -gt 0) {
    # Validate / normalize the supplied URLs
    $sites = $SiteUrls | ForEach-Object {
        [PSCustomObject]@{
            Url   = $_.TrimEnd('/')
            Title = $_.TrimEnd('/').Split('/')[-1]
        }
    }
    Write-OK "$($sites.Count) site(s) specified."
} else {
    Write-Step "Retrieving all SharePoint sites (this may take a moment) …"
    $sites = Get-PnPTenantSite -IncludeOneDriveSites:$false |
             Where-Object {
                 $_.Template -notlike '*SPSPERS*' -and
                 $_.Template -notlike '*REDIRECT*'
             } |
             ForEach-Object {
        [PSCustomObject]@{
            Url   = $_.Url.TrimEnd('/')
            Title = $_.Title
        }
    }
    Write-OK "Found $($sites.Count) SharePoint site(s)."
}
#endregion
#region ── Data collection ────────────────────────────────────────────────────
Write-Header "Collecting Search Insights"
$columnOrder = @(
    'ExportDate','ReportingPeriod','Source','SiteTitle','SiteUrl',
    'InsightType','Date','QueryText',
    'QueryCount','AbandonedCount','NoResultsCount','ClickCount'
)

if (-not (Test-Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
}

$timestamp       = Get-Date -Format 'yyyyMMdd_HHmmss'
$csvFile         = Join-Path $OutputFolder "SPO_SearchInsights_${periodLabel}_${timestamp}.csv"
$bufferFlushSize = 500
$rowBuffer       = [System.Collections.Generic.List[PSCustomObject]]::new()
$seenRowKeys     = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$runExportDate   = Get-Date -Format 'yyyy-MM-dd HH:mm'
$totalRows       = 0
$insightCounts   = @{}
$siteCounts      = @{}

function Flush-RowBuffer {
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
            Flush-RowBuffer
        }
    }
}
# ── Helper: call SharePoint REST analytics endpoint ──────────────────────────
function Invoke-SiteSearchAnalytics {
    param(
        [string]$SiteUrl,
        [string]$AnalyticsEndpoint,   # relative, e.g. "GetTopQueries"
        [string]$StartDate,
        [string]$EndDate
    )
    $apiPath = "$SiteUrl/_api/search/analytics/$AnalyticsEndpoint" +
               "?startdate='$StartDate'&enddate='$EndDate'&rowlimit=100"
    try {
        $resp = Invoke-PnPSPRestMethod -Method Get -Url $apiPath
        return $resp
    } catch {
        $statusCode = $null
        $responseBody = $null
        if ($_.Exception.PSObject.Properties['Response'] -and $_.Exception.Response) {
            $statusCode = try { [int]$_.Exception.Response.StatusCode } catch { $null }
            try {
                $stream = $_.Exception.Response.GetResponseStream()
                if ($stream) {
                    $reader = [System.IO.StreamReader]::new($stream)
                    $responseBody = $reader.ReadToEnd()
                    $reader.Dispose()
                }
            } catch {
                $responseBody = $null
            }
        }

        Write-Warn "REST analytics call failed for $SiteUrl [$AnalyticsEndpoint]. Status: $statusCode. Error: $($_.Exception.Message)"
        if ($responseBody) {
            Write-Warn "Response body: $responseBody"
        }

        return $null
    }
}
# ── Helper: parse SharePoint REST analytics response ─────────────────────────
function ConvertFrom-SPOAnalyticsResponse {
    param(
        [object]$Response,
        [string]$SiteTitle,
        [string]$SiteUrl,
        [string]$InsightType,
        [string]$PeriodLabel
    )
    if (-not $Response) { return }
    $items = $null
    if ($Response.PSObject.Properties['value'])   { $items = $Response.value }
    elseif ($Response -is [array])                { $items = $Response }
    if (-not $items) { return }
    foreach ($item in $items) {
        [PSCustomObject]@{
            ExportDate      = $runExportDate
            ReportingPeriod = $PeriodLabel
            Source          = 'SPO_REST'
            SiteTitle       = $SiteTitle
            SiteUrl         = $SiteUrl
            InsightType     = $InsightType
            Date            = if ($item.PSObject.Properties['Date'])         { $item.Date }         else { '' }
            QueryText       = if ($item.PSObject.Properties['Query'])        { $item.Query }
                              elseif ($item.PSObject.Properties['QueryText']){ $item.QueryText }    else { '' }
            QueryCount      = if ($item.PSObject.Properties['Count'])        { $item.Count }
                              elseif ($item.PSObject.Properties['Hits'])     { $item.Hits }         else { '' }
            AbandonedCount  = if ($item.PSObject.Properties['Abandoned'])    { $item.Abandoned }    else { '' }
            NoResultsCount  = if ($item.PSObject.Properties['NoResults'])    { $item.NoResults }    else { '' }
            ClickCount      = if ($item.PSObject.Properties['Clicks'])       { $item.Clicks }       else { '' }
        }
    }
}
# ── Per-site SharePoint REST analytics (supplemental) ─────────────────────────
$siteIndex = 0
foreach ($site in $sites) {
    $siteIndex++
    $pct = [int](($siteIndex / $sites.Count) * 100)
    Write-Progress -Activity "Querying site analytics" `
                   -Status "$siteIndex of $($sites.Count): $($site.Title)" `
                   -PercentComplete $pct
    $spoInsights = @(
        @{ Type = 'TopQueries';       Endpoint = 'GetTopQueries' }
        @{ Type = 'NoResultQueries';  Endpoint = 'GetNoResultQueries' }
        @{ Type = 'AbandonedQueries'; Endpoint = 'GetAbandonedQueries' }
    )
    foreach ($insight in $spoInsights) {
        $resp = Invoke-SiteSearchAnalytics `
                    -SiteUrl       $site.Url `
                    -AnalyticsEndpoint $insight.Endpoint `
                    -StartDate     $startDateStr `
                    -EndDate       $endDateStr
        $parsed = ConvertFrom-SPOAnalyticsResponse `
                    -Response    $resp `
                    -SiteTitle   $site.Title `
                    -SiteUrl     $site.Url `
                    -InsightType $insight.Type `
                    -PeriodLabel $periodLabel
        if ($parsed) {
            Add-ResultRows -Rows $parsed
        }
    }
    if ($HeartbeatIntervalSites -gt 0 -and
        (($siteIndex % $HeartbeatIntervalSites) -eq 0 -or $siteIndex -eq $sites.Count)) {
        Write-Step "Heartbeat: processed $siteIndex/$($sites.Count) sites ($pct%). Rows written so far: $totalRows"
    }
}
Write-Progress -Activity "Querying site analytics" -Completed
#endregion
#region ── Fallback: Graph site usage report ──────────────────────────────────
# If no search-specific rows were collected, fall back to the GA usage report
# which at minimum shows page views and site activity per site.
if ($totalRows -eq 0) {
    Write-Warn "No search-insight data returned from Graph or REST APIs."
    Write-Step "Falling back to SharePoint Site Usage report (D30 / D180) …"
    $usagePeriod = if ($selectedReportPeriod -eq 'Last28Days') { 'D30' } else { 'D180' }
    try {
        $usageUrl  = "v1.0/reports/getSharePointSiteUsageDetail(period='$usagePeriod')"
        $usageCsv  = Invoke-PnPGraphMethod -Url $usageUrl -Method Get -Raw
        $usageRows = $usageCsv | ConvertFrom-Csv
        foreach ($row in $usageRows) {
            $siteUrl  = ($row.'Site URL' -replace '/$','')
            $included = (-not $SiteUrls -or $SiteUrls.Count -eq 0) -or ($SiteUrls | Where-Object { $_.TrimEnd('/') -eq $siteUrl })
            if (-not $included) { continue }
            Add-ResultRows -Rows @([PSCustomObject]@{
                ExportDate      = $runExportDate
                ReportingPeriod = $periodLabel
                Source          = 'GRAPH_USAGE'
                SiteTitle       = $row.'Site Name'
                SiteUrl         = $siteUrl
                InsightType     = 'SiteUsageSummary'
                Date            = $row.'Report Refresh Date'
                QueryText       = ''
                QueryCount      = $row.'Visited Page Count'
                AbandonedCount  = ''
                NoResultsCount  = ''
                ClickCount      = $row.'Page View Count'
            })
        }
        Write-OK "Site usage report: $totalRows row(s) collected."
    } catch {
        Write-Fail "Site usage report also failed: $_"
    }
}
#endregion
#region ── Export ─────────────────────────────────────────────────────────────
Write-Header "Exporting Results"
if ($totalRows -eq 0) {
    Write-Warn "No data was collected.  The CSV will not be created."
    Write-Host ""
    Write-Host "  Possible reasons:" -ForegroundColor Gray
    Write-Host "  • Your account lacks Reports.Read.All Graph permission." -ForegroundColor Gray
    Write-Host "  • Search Insights haven't been enabled on these sites." -ForegroundColor Gray
    Write-Host "  • The selected period has no recorded search activity." -ForegroundColor Gray
} else {
    Flush-RowBuffer
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
    Source           Data source (GRAPH, SPO_REST, GRAPH_USAGE)
  SiteTitle        Display name of the site
  SiteUrl          Absolute URL of the site
  InsightType      TopQueries | NoResultQueries | AbandonedQueries | SiteUsageSummary
  Date             Date the metric relates to (daily / monthly / summary)
  QueryText        The search query string (blank for summary rows)
  QueryCount       Number of times the query was executed (or visited page count for summary)
  AbandonedCount   Queries started but not completed
  NoResultsCount   Queries that returned zero results
  ClickCount       Number of result clicks (or page view count for summary)
#>
#endregion