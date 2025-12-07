# Created by Daniel Glenn 10/28/2025, revised December 6th, 2025
# Repository: https://github.com/danielglenn/PowerShell-for-Microsoft-365

# Prerequisites:
# 1. Install the PnP.PowerShell module if not already installed:
#    Install-Module -Name PnP.PowerShell -Scope CurrentUser
# 2. Create an Entra ID App Registration with certificate-based authentication and appropriate SharePoint permissions.
#    Grant the app registration the necessary SharePoint permissions (e.g., Sites.Read.All) via the Entra ID portal.
# 3. Ensure you have a certificate with a private key installed in the CurrentUser\My store on your computer and note its thumbprint.
# 4. Edit the $sites array below to include the full URLs of the sites you want to scan. *Will update this in future to read from a file and/or from user input.*

# Usage:
# To call this script, use the following syntax, replacing the parameters with your own values:
#  .\ListAllSharePointSiteFiles.ps1 -ClientId "your-app-id" -TenantId "your-tenant-id" -Thumbprint "ABC123DEF456..." -exportFolder "C:\Exports"

[CmdletBinding()]
param(
    
    [Parameter(Mandatory = $true)]
    [string]$ClientId,  # Entra ID App (Client) ID
    
    [Parameter(Mandatory = $true)]
    [string]$TenantId,  # Entra ID Tenant ID (GUID or name.onmicrosoft.com)
    
    [Parameter(Mandatory = $true)]
    [string]$Thumbprint  # Certificate thumbprint from CurrentUser\My store

    [Parameter(Mandatory = $true)]
    [string]$exportFolder  # FOLDER of the CSV to write to, such as "C:\Exports"
)
# Define the list of sites to process, each URL should be the full URL to the site
$sites = @(
	"Site1FullURL",
	"Site2FullURL",
    "Site3FullURL"

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

# initializing the CSV file location
If (-not (Test-Path $exportFolder)) {
	New-Item -ItemType Directory -Path $exportFolder | Out-Null
}
# loop through each site
foreach ($site in $sites) {
    Write-Host "Connecting to $site..."
# Connect to the site using the certificate - the certificate should be in the Current User store locally
	Connect-PnPOnline -URL "$site" -clientID "$clientID" -tenant "$TenantId" -thumbprint "$thumbprint" 

    # Generate a safe filename based on site name
    $siteName = ($site.Split("/")[-1]).Replace(" ", "_")
    $csvPath = Join-Path $exportFolder "$siteName-files.csv"

    # Initialize in-memory buffer
    $results = @()

    # Get all document libraries (BaseTemplate 101) that are not hidden
    $libraries = Get-PnPList | Where-Object { $_.BaseTemplate -eq 101 -and $_.Hidden -eq $false }

    foreach ($lib in $libraries) {
        Write-Host "Scanning library: $($lib.Title)"

        # Get all items in the library
        $items = Get-PnPListItem -List $lib.Title -PageSize 1000 -Fields "FileRef","FileLeafRef","FSObjType"

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
                    SiteUrl  = $site
                    Library  = $lib.Title
                    FileName = $item.FieldValues["FileLeafRef"]
                    FileSize = $fileSizeMB
                    FileUrl  = $fileUrl
                    LastModified = $lastModified
                }
            }
        }
    }

    # Write all results for this site in one go
    $results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
    Write-Host "Exported $($results.Count) files to $csvPath"
}
