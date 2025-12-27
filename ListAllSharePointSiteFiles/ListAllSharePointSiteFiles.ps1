# Created by Daniel Glenn 10/28/2025, revised December 27th, 2025
# Repository: https://github.com/danielglenn/PowerShell-for-Microsoft-365

# Prerequisites:
# 1. Install the PnP.PowerShell module if not already installed:
#    Install-Module -Name PnP.PowerShell -Scope CurrentUser
# 2. Create an Entra ID App Registration with certificate-based authentication and appropriate SharePoint permissions.
#    Grant the app registration the necessary SharePoint permissions (e.g., Sites.Read.All) via the Entra ID portal.
# 3. Ensure you have a certificate with a private key installed in the CurrentUser\My store on your computer and note its thumbprint.
# 4. Provide a CSV file with a 'SiteUrl' column listing the full URLs of the sites you want to scan.

# Usage:
# To call this script, use the following syntax, replacing the parameters with your own values:
#  .\ListAllSharePointSiteFiles.ps1 -ClientId "your-app-id" -TenantId "your-tenant-id" -Thumbprint "ABC123DEF456..." -SitesCsvPath "C:\Exports\sites.csv" -exportFolder "C:\Exports"

[CmdletBinding()]
param(
    
    [Parameter(Mandatory = $true)]
    [string]$ClientId,  # Entra ID App (Client) ID
    
    [Parameter(Mandatory = $true)]
    [string]$TenantId,  # Entra ID Tenant ID (GUID or name.onmicrosoft.com)
    
    [Parameter(Mandatory = $true)]
    [string]$Thumbprint,  # Certificate thumbprint from CurrentUser\My store

    [Parameter(Mandatory = $true)]
    [string]$SitesCsvPath,  # Path to CSV containing a 'SiteUrl' column

    [Parameter(Mandatory = $true)]
    [string]$exportFolder  # FOLDER of the CSV to write to, such as "C:\Exports"
)
# Load site list from CSV
if (-not (Test-Path $SitesCsvPath)) {
    Write-Error "Sites CSV not found at '$SitesCsvPath'."
    return
}

$csvRows = Import-Csv -Path $SitesCsvPath -ErrorAction Stop

if (-not ($csvRows | Get-Member -Name SiteUrl -MemberType NoteProperty)) {
    Write-Error "Sites CSV must contain a 'SiteUrl' column."
    return
}

$sites = $csvRows |
    ForEach-Object { $_.SiteUrl } |
    Where-Object { $_ -and $_.Trim() } |
    ForEach-Object { $_.Trim() } |
    Sort-Object -Unique

if (-not $sites) {
    Write-Error "No valid SiteUrl entries found in '$SitesCsvPath'."
    return
}

Write-Host "Loaded $($sites.Count) site(s) from $SitesCsvPath" -ForegroundColor Cyan

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

# initializing the CSV file location
If (-not (Test-Path $exportFolder)) {
	New-Item -ItemType Directory -Path $exportFolder | Out-Null
}
# loop through each site
foreach ($site in $sites) {
    if ([string]::IsNullOrWhiteSpace($site)) {
        Write-Warning "Skipping blank SiteUrl entry in CSV."
        continue
    }

    $siteCandidate = $site.Trim()
    $uri = $null
    if (-not [Uri]::TryCreate($siteCandidate, [UriKind]::Absolute, [ref]$uri) -or $uri.Scheme -ne "https") {
        Write-Warning "Skipping invalid site URL '$siteCandidate' from CSV."
        continue
    }

    $siteNormalized = $uri.AbsoluteUri.TrimEnd("/")
    Write-Host "Connecting to $siteNormalized..."
    try {
        # Connect to the site using the certificate in the Current User store
        Connect-PnPOnline -Url $siteNormalized -ClientId $ClientId -Tenant $TenantId -Thumbprint $Thumbprint -ErrorAction Stop
        # Validate the connection by attempting to get the web properties
        $web = Get-PnPWeb -ErrorAction Stop
    }
    catch {
        Write-Warning "Failed to connect to '$siteNormalized'. Skipping. Error: $($_.Exception.Message)"
        continue
    }

    # Generate a safe filename based on site name
    $siteSegment = $siteNormalized.Split("/")[-1]
    if (-not $siteSegment) { $siteSegment = $siteNormalized.Split("/")[-2] }
    $siteSegment = $siteSegment.Trim()

    # Strip characters invalid for filenames and normalize underscores
    $siteName = [Regex]::Replace($siteSegment, '[\\\/:*?"<>|]+', "_").Trim("_")
    if (-not $siteName) {
        $siteName = "site-" + ([Guid]::NewGuid().ToString("N").Substring(8))
    }

    $csvPath = Join-Path $exportFolder "$siteName-files.csv"

    # Initialize in-memory buffer
    $results = @()

    try {
        # Get all document libraries (BaseTemplate 101) that are not hidden
        $libraries = Get-PnPList -ErrorAction Stop | Where-Object { $_.BaseTemplate -eq 101 -and $_.Hidden -eq $false }

        foreach ($lib in $libraries) {
            Write-Host "Scanning library: $($lib.Title)"

            try {
                # Get all items in the library
                $items = Get-PnPListItem -List $lib.Title -PageSize 1000 -Fields "FileRef","FileLeafRef","FSObjType" -ErrorAction Stop

                foreach ($item in $items) {
                    # Check if the item is a file (FSObjType 0) - folders are FSObjType 1
                    if ($item.FieldValues["FileRef"] -and $item.FieldValues["FSObjType"] -eq 0) {
                        $fileRef = [string]$item.FieldValues["FileRef"]
                        $fileUrl = "https://$($TenantId.Split('.')[0]).sharepoint.com$fileRef"
                        # Get the File object for this item
                        $file = Get-PnPProperty -ClientObject $item -Property File
                        # Retrieve file size in MB
                        $fileSizeMB = [math]::Round($file.Length / 1MB, 2)
                        # Retrieve last modified date
                        $lastModified = $file.TimeLastModified
                        $results += [PSCustomObject]@{
                            SiteUrl  = $siteNormalized
                            Library  = $lib.Title
                            FileName = $item.FieldValues["FileLeafRef"]
                            FileSize = $fileSizeMB
                            FileUrl  = $fileUrl
                            LastModified = $lastModified
                        }
                    }
                }
            }
            catch {
                Write-Warning "Error scanning library '$($lib.Title)' on '$siteNormalized'. Skipping library. Error: $($_.Exception.Message)"
                continue
            }
        }
    }
    catch {
        Write-Warning "Failed to retrieve libraries from '$siteNormalized'. Skipping site. Error: $($_.Exception.Message)"
        continue
    }

    # Write all results for this site in one go
    $results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
    Write-Host "Exported $($results.Count) files to $csvPath"
}
