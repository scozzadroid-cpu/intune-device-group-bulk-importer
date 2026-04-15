# Convert device serial numbers to Azure AD Object IDs for Intune Group Import
# Uses Microsoft.Graph module with multiple search strategies

$SerialsFile = "C:\TEMP\serials.txt"
$TemplateCSV = "C:\TEMP\GroupImportMembersTemplate.csv"
$OutputCSV = "C:\TEMP\IntuneGroupImport.csv"

# Install Microsoft.Graph module if not present
if (!(Get-Module -ListAvailable -Name Microsoft.Graph.Identity.DirectoryManagement)) {
    Write-Host "Installing Microsoft.Graph.Identity.DirectoryManagement module"
    Install-Module -Name Microsoft.Graph.Identity.DirectoryManagement -Force -AllowClobber
}

# Connect to Microsoft Graph
Write-Host "Connecting to Microsoft Graph"
Connect-MgGraph -Scopes "Device.Read.All"

# Verify input files exist
if (!(Test-Path $SerialsFile)) {
    Write-Error "Serials file not found at $SerialsFile"
    exit 1
}

if (!(Test-Path $TemplateCSV)) {
    Write-Host "Template CSV not found - creating it automatically at $TemplateCSV" -ForegroundColor Yellow
    @(
        "Member object ID or user principal name [memberObjectIdOrUpn] Required",
        "Example: 9832aad8-e4fe-496b-a604-95c6eF01ae75"
    ) | Out-File -FilePath $TemplateCSV -Encoding UTF8
    Write-Host "Template CSV created." -ForegroundColor Green
}

# Read serials and remove duplicates
$DeviceSerials = Get-Content $SerialsFile | Where-Object { $_ -and $_ -notmatch "^\s*$" } | Select-Object -Unique
Write-Host "Loaded $($DeviceSerials.Count) unique serials"

# Copy template and prepare output
Copy-Item $TemplateCSV $OutputCSV -Force
$TemplateContent = Get-Content $TemplateCSV

# Find header end index
$HeaderEndIndex = -1
$HeaderPatterns = @("\[memberObjectIdOrUpn\]", "Member object ID", "memberObjectIdOrUpn", "Object ID", "ObjectId")

for ($i = 0; $i -lt $TemplateContent.Count; $i++) {
    foreach ($Pattern in $HeaderPatterns) {
        if ($TemplateContent[$i] -match $Pattern) {
            $HeaderEndIndex = $i
            break
        }
    }
    if ($HeaderEndIndex -ne -1) { break }
}

if ($HeaderEndIndex -eq -1) {
    Write-Error "Invalid template format"
    exit 1
}

$HeaderContent = $TemplateContent[0..$HeaderEndIndex]
$HeaderContent | Out-File -FilePath $OutputCSV -Encoding UTF8

# Get ALL devices once via Microsoft Graph to improve performance
# Uses Invoke-MgGraphRequest with manual pagination — more reliable than
# Get-MgDevice -All which throws AggregateException on some SDK versions.
Write-Host "Loading all devices from Microsoft Graph (this may take a moment)"
$AllDevices = [System.Collections.Generic.List[object]]::new()
$NextUri = "https://graph.microsoft.com/v1.0/devices?`$select=displayName,id"
do {
    $Response = Invoke-MgGraphRequest -Uri $NextUri -Method GET
    foreach ($Dev in $Response.value) { $AllDevices.Add($Dev) }
    $NextUri = $Response['@odata.nextLink']
} while ($NextUri)
Write-Host "Loaded $($AllDevices.Count) devices from Microsoft Graph"
Write-Host ""

# Collections for results
$FoundObjectIDs = @()
$ProcessedCount = 0
$NotFoundCount = 0
$MethodStats = @{
    "Exact" = 0
    "StartsWith" = 0
    "Contains" = 0
    "CaseInsensitive" = 0
}

Write-Host "Starting device search with multiple fallback strategies"
Write-Host ""

# Process each serial number
foreach ($SerialNumber in $DeviceSerials) {
    
    if ([string]::IsNullOrWhiteSpace($SerialNumber)) {
        continue
    }
    
    $Device = $null
    $Method = $null
    
    # Strategy 1 - Exact match on DisplayName
    $Device = $AllDevices | Where-Object { $_.displayName -eq $SerialNumber } | Select-Object -First 1
    if ($Device) { 
        $Method = "Exact"
    }
    
    # Strategy 2 - DisplayName starts with serial
    if (!$Device) {
        $Device = $AllDevices | Where-Object { $_.displayName -like "$SerialNumber*" } | Select-Object -First 1
        if ($Device) { 
            $Method = "StartsWith"
        }
    }
    
    # Strategy 3 - DisplayName contains serial
    if (!$Device) {
        $Device = $AllDevices | Where-Object { $_.displayName -like "*$SerialNumber*" } | Select-Object -First 1
        if ($Device) { 
            $Method = "Contains"
        }
    }
    
    # Strategy 4 - Case-insensitive fuzzy search
    if (!$Device) {
        $Device = $AllDevices | Where-Object { $_.displayName -match [regex]::Escape($SerialNumber) } | Select-Object -First 1
        if ($Device) { 
            $Method = "CaseInsensitive"
        }
    }
    
    if ($Device) {
        $FoundObjectIDs += $Device.id
        Write-Host "[$Method] $SerialNumber -> $($Device.displayName) -> $($Device.id)"
        $ProcessedCount++
        $MethodStats[$Method]++
    }
    else {
        Write-Warning "NOT FOUND $SerialNumber"
        $NotFoundCount++
    }
}

Write-Host ""
Write-Host "Processing complete"

# Remove duplicate Object IDs
$UniqueObjectIDs = $FoundObjectIDs | Select-Object -Unique
$RemovedDuplicates = $FoundObjectIDs.Count - $UniqueObjectIDs.Count

if ($RemovedDuplicates -gt 0) {
    Write-Host "Removed $RemovedDuplicates duplicate Object IDs"
}

# Append unique Object IDs to CSV
$UniqueObjectIDs | Out-File -FilePath $OutputCSV -Append -Encoding UTF8

# Summary
Write-Host ""
Write-Host "SUMMARY"
Write-Host "Total serials processed $($DeviceSerials.Count)"
Write-Host "Devices found $ProcessedCount"
Write-Host "  - Exact match $($MethodStats['Exact'])"
Write-Host "  - Starts with $($MethodStats['StartsWith'])"
Write-Host "  - Contains $($MethodStats['Contains'])"
Write-Host "  - Case insensitive $($MethodStats['CaseInsensitive'])"
Write-Host "Devices not found $NotFoundCount"
Write-Host "Unique Object IDs added $($UniqueObjectIDs.Count)"
Write-Host ""
Write-Host "Output file ready $OutputCSV"
Write-Host ""
Write-Host "NEXT STEPS"
Write-Host "1. Open Azure Portal > Groups"
Write-Host "2. Select target group > Members > Bulk operations > Import members"
Write-Host "3. Upload $OutputCSV"
Write-Host "4. Check bulk operation results for any errors"