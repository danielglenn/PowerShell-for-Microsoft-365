# Created by Daniel Glenn, December 27th, 2025
# Repository: https://github.com/danielglenn/PowerShell-for-Microsoft-365

# Prerequisites:
# 1. Install the PnP.PowerShell module if not already installed:
#    Install-Module -Name PnP.PowerShell -Scope CurrentUser
# 2. Create an Entra ID App Registration with certificate-based authentication and appropriate SharePoint permissions.
#    Grant the app registration the necessary SharePoint permissions (e.g., Sites.FullControl.All) via the Entra ID portal.
# 3. Ensure you have a certificate with a private key installed in the CurrentUser\My store on your computer and note its thumbprint.
# 4. Provide either a CSV file with a 'SiteUrl' column or an array of site URLs.

# Usage:
# Using CSV file:
#  .\ListFilesWithUniquePermissions.ps1 -ClientId "your-app-id" -TenantId "your-tenant-id" -Thumbprint "ABC123DEF456..." -SitesCsvPath "C:\Exports\sites.csv" -exportFolder "C:\Exports"
# Using array of URLs:
#  .\ListFilesWithUniquePermissions.ps1 -ClientId "your-app-id" -TenantId "your-tenant-id" -Thumbprint "ABC123DEF456..." -SiteUrls @("https://contoso.sharepoint.com/sites/site1","https://contoso.sharepoint.com/sites/site2") -exportFolder "C:\Exports"

[CmdletBinding(DefaultParameterSetName = 'CsvFile')]
param(
    [Parameter(Mandatory = $true)]
    [string]$ClientId,
    
    [Parameter(Mandatory = $true)]
    [string]$TenantId,
    
    [Parameter(Mandatory = $true)]
    [string]$Thumbprint,

    [Parameter(Mandatory = $true, ParameterSetName = 'CsvFile')]
    [string]$SitesCsvPath,

    [Parameter(Mandatory = $true, ParameterSetName = 'DirectUrls')]
    [string[]]$SiteUrls,

    [Parameter(Mandatory = $true)]
    [string]$exportFolder
)

function Format-PermissionsArray {
    param([Parameter(Mandatory=$true)]$RoleAssignments)

    $assignmentsArray = @($RoleAssignments)
    if (-not $assignmentsArray -or $assignmentsArray.Count -eq 0) {
        return @("No permissions assigned")
    }

    $permissionsList = @()
    foreach ($assignment in $assignmentsArray) {
        try {
            Get-PnPProperty -ClientObject $assignment -Property Member, RoleDefinitionBindings -ErrorAction Stop | Out-Null

            $member = $assignment.Member
            if (-not $member) { continue }

            Get-PnPProperty -ClientObject $member -Property LoginName, Title -ErrorAction Stop | Out-Null
            $loginName = if ($member.LoginName) { $member.LoginName } else { $member.Title }
            if (-not $loginName) { continue }

            $roleBindings = @($assignment.RoleDefinitionBindings)
            if ($roleBindings.Count -gt 0) {
                $roles = @()
                foreach ($role in $roleBindings) {
                    Get-PnPProperty -ClientObject $role -Property Name -ErrorAction Stop | Out-Null
                    if ($role.Name) { $roles += $role.Name }
                }
                $permissionsList += "$loginName (" + ($roles -join ', ') + ")"
            } else {
                $permissionsList += "$loginName (No Role)"
            }
        } catch {
            continue
        }
    }

    if ($permissionsList.Count -eq 0) {
        return @("No permissions assigned")
    }

    return $permissionsList
}

function Format-Permissions {
    param([Parameter(Mandatory=$true)]$RoleAssignments)
    $permArray = Format-PermissionsArray -RoleAssignments $RoleAssignments
    return [string]::Join("; ", $permArray)
}

function Write-LibraryPermissionsFile {
    param(
        [string]$SiteUrl,
        [string]$LibraryTitle,
        $RoleAssignments,
        [string]$SiteName,
        [string]$ExportFolder
    )

    $libNameSegment = [Regex]::Replace($LibraryTitle, '[\\/:*?"<>|]+', "_").Trim("_")
    if (-not $libNameSegment) { $libNameSegment = "library-" + ([Guid]::NewGuid().ToString("N").Substring(8)) }

    $txtPath = Join-Path $ExportFolder "$SiteName-$libNameSegment-perms.txt"

    $permArray = Format-PermissionsArray -RoleAssignments $RoleAssignments

    $content = @()
    $content += "Site: $SiteUrl"
    $content += "Library: $LibraryTitle"
    $content += "Permissions:"
    foreach ($perm in $permArray) {
        $content += "  - $perm"
    }

    Set-Content -Path $txtPath -Value $content -Encoding UTF8
}

# Load site list
if ($PSCmdlet.ParameterSetName -eq 'CsvFile') {
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

    Write-Host "Loaded $($sites.Count) site(s) from CSV file" -ForegroundColor Cyan
}
else {
    $sites = $SiteUrls |
        Where-Object { $_ -and $_.Trim() } |
        ForEach-Object { $_.Trim() } |
        Sort-Object -Unique

    if (-not $sites) {
        Write-Error "No valid site URLs provided in -SiteUrls parameter."
        return
    }

    Write-Host "Loaded $($sites.Count) site(s) from parameter" -ForegroundColor Cyan
}

Import-Module PnP.PowerShell -ErrorAction Stop

Write-Host "Validating certificate thumbprint in CurrentUser\My store..." -ForegroundColor Cyan
$cert = Get-ChildItem -Path Cert:\CurrentUser\My -ErrorAction SilentlyContinue | Where-Object { $_.Thumbprint -eq $Thumbprint }
if (-not $cert) {
    Write-Error "Certificate with thumbprint '$Thumbprint' not found in CurrentUser\My store."
    return
}
Write-Host "Certificate validated: $($cert.Subject)" -ForegroundColor Green

if (-not (Test-Path $exportFolder)) {
    New-Item -ItemType Directory -Path $exportFolder | Out-Null
}

foreach ($site in $sites) {
    if ([string]::IsNullOrWhiteSpace($site)) {
        Write-Warning "Skipping blank SiteUrl entry."
        continue
    }

    $siteCandidate = $site.Trim()
    $uri = $null
    if (-not [Uri]::TryCreate($siteCandidate, [UriKind]::Absolute, [ref]$uri) -or $uri.Scheme -ne "https") {
        Write-Warning "Skipping invalid site URL '$siteCandidate'."
        continue
    }

    $siteNormalized = $uri.AbsoluteUri.TrimEnd("/")
    Write-Host "Connecting to $siteNormalized..." -ForegroundColor Cyan
    try {
        Connect-PnPOnline -Url $siteNormalized -ClientId $ClientId -Tenant $TenantId -Thumbprint $Thumbprint -ErrorAction Stop
        $web = Get-PnPWeb -ErrorAction Stop
    }
    catch {
        Write-Warning "Failed to connect to '$siteNormalized'. Skipping. Error: $($_.Exception.Message)"
        continue
    }

    $siteSegment = $siteNormalized.Split("/")[-1]
    if (-not $siteSegment) { $siteSegment = $siteNormalized.Split("/")[-2] }
    $siteSegment = $siteSegment.Trim()

    $siteName = [Regex]::Replace($siteSegment, '[\\\/:*?"<>|]+', "_").Trim("_")
    if (-not $siteName) {
        $siteName = "site-" + ([Guid]::NewGuid().ToString("N").Substring(8))
    }

    $csvPath = Join-Path $exportFolder "$siteName-unique-permissions.csv"
    $results = @()

    try {
        $libraries = Get-PnPList -ErrorAction Stop | Where-Object { $_.BaseTemplate -eq 101 -and $_.Hidden -eq $false }

        foreach ($lib in $libraries) {
            Write-Host "  Scanning library: $($lib.Title)" -ForegroundColor Yellow

            try {
                $items = Get-PnPListItem -List $lib.Title -PageSize 1000 -ErrorAction Stop
                $libHasUniqueFiles = $false

                foreach ($item in $items) {
                    $itemType = $item.FieldValues["FSObjType"]
                    if ($itemType -eq 0) {
                        try {
                            $hasUniquePermissions = Get-PnPProperty -ClientObject $item -Property HasUniqueRoleAssignments
                            if ($hasUniquePermissions) {
                                # Mark that this library has at least one file with unique permissions
                                $libHasUniqueFiles = $true
                                
                                # Get the role assignments for this file
                                Get-PnPProperty -ClientObject $item -Property RoleAssignments | Out-Null
                                $filePermissions = [string](Format-Permissions -RoleAssignments $item.RoleAssignments)

                                $fileRef = [string]$item.FieldValues["FileRef"]
                                $fileUrl = "https://$($TenantId.Split('.')[0]).sharepoint.com$fileRef"

                                $results += [PSCustomObject]@{
                                    SiteUrl             = $siteNormalized
                                    Library             = $lib.Title
                                    FileName            = $item.FieldValues["FileLeafRef"]
                                    FileUrl             = $fileUrl
                                    FilePermissions     = $filePermissions
                                    HasUniquePermissions = $true
                                }
                            }
                        }
                        catch {
                            Write-Warning "    Error checking permissions for file '$($item.FieldValues["FileLeafRef"])'. Error: $($_.Exception.Message)"
                        }
                    }
                }

                # Only create library permissions file if library has files with unique permissions
                if ($libHasUniqueFiles) {
                    $libWithPerms = Get-PnPList -Identity $lib.Title -ErrorAction Stop
                    Get-PnPProperty -ClientObject $libWithPerms -Property RoleAssignments | Out-Null
                    Write-LibraryPermissionsFile -SiteUrl $siteNormalized -LibraryTitle $lib.Title -RoleAssignments $libWithPerms.RoleAssignments -SiteName $siteName -ExportFolder $exportFolder
                }
            }
            catch {
                Write-Warning "  Error scanning library '$($lib.Title)' on '$siteNormalized'. Skipping library. Error: $($_.Exception.Message)"
                continue
            }
        }
    }
    catch {
        Write-Warning "Failed to retrieve libraries from '$siteNormalized'. Skipping site. Error: $($_.Exception.Message)"
        continue
    }

    if ($results.Count -gt 0) {
        $results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
        Write-Host "Exported $($results.Count) file(s) with unique permissions to $csvPath" -ForegroundColor Green
    }
    else {
        Write-Host "No files with unique permissions found on this site." -ForegroundColor Gray
    }
}

Write-Host "`nProcessing complete." -ForegroundColor Cyan
