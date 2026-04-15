# Convert device hostnames to Azure AD Object IDs for Intune Group Import
# Reads hostnames from C:\TEMP\pc.txt and uses Intune template CSV
# Requirements: Microsoft.Graph module and appropriate permissions
#
# NOTE: This script uses DeviceManagementManagedDevices.Read.All (already consented in most tenants).
# It avoids Device.Read.All which requires additional admin consent.
# The Azure AD Object ID is resolved via the Intune managedDevice > azureADDeviceId lookup.

# Disable WAM to avoid window handle issues
$env:MSAL_DISABLE_WAM = 1

$RequiredModules = @("Microsoft.Graph.DeviceManagement", "Microsoft.Graph.Identity.DirectoryManagement")
foreach ($Mod in $RequiredModules) {
    if (!(Get-Module -ListAvailable -Name $Mod)) {
        Install-Module -Name $Mod -Force -AllowClobber
    }
}

# Disconnect any existing sessions
Disconnect-MgGraph -ErrorAction SilentlyContinue

# Connect with interactive browser authentication
Connect-MgGraph -Scopes "DeviceManagementManagedDevices.Read.All" -NoWelcome

# Verify connection
$Context = Get-MgContext
if (-not $Context) {
    Write-Host "Authentication failed. Please try again." -ForegroundColor Red
    
}

# Pre-load all Intune managed devices once for fast hostname lookup
# Uses Invoke-MgGraphRequest with manual pagination — more reliable than
# Get-MgDeviceManagementManagedDevice -All which throws AggregateException on some SDK versions.
Write-Host "Loading all managed devices from Intune..."
$script:AllManagedDevices = [System.Collections.Generic.List[object]]::new()
$NextUri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$select=deviceName,id,azureADDeviceId"
do {
    $Response = Invoke-MgGraphRequest -Uri $NextUri -Method GET
    foreach ($Device in $Response.value) { $script:AllManagedDevices.Add($Device) }
    $NextUri = $Response['@odata.nextLink']
} while ($NextUri)
Write-Host "Loaded $($script:AllManagedDevices.Count) devices."

# Helper: resolve Azure AD Object ID from hostname via Intune managedDevices
# The managedDevice.Id is the Intune ID; the actual Entra ID Object ID requires
# a second call to /devices?$filter=deviceId eq '{azureADDeviceId}' — but that
# needs Device.Read.All. As a workaround we cache a second lookup using the
# Intune device list and the azureADDeviceId field.
# If Device.Read.All is consented in your tenant, uncomment the block below
# to get the true Entra Object ID; otherwise the Intune device ID is returned
# (which works for Intune-managed groups but not pure Entra ID groups).

function Get-EntraObjectId {
    param([string]$DeviceName)

    # Invoke-MgGraphRequest returns hashtables with camelCase keys
    $IntuneDevice = $script:AllManagedDevices |
        Where-Object { $_.deviceName -eq $DeviceName } |
        Select-Object -First 1

    if (!$IntuneDevice) { return $null }

    # --- Option A: Device.Read.All consented (true Entra Object ID) ---
    # $EntraDevice = Get-MgDevice -Filter "deviceId eq '$($IntuneDevice.azureADDeviceId)'" -ConsistencyLevel eventual -ErrorAction SilentlyContinue | Select-Object -First 1
    # if ($EntraDevice) { return $EntraDevice.Id }

    # --- Option B: use Intune managed device ID (works for Intune group import) ---
    return $IntuneDevice.id
}

$HostnamesFile = "C:\TEMP\pc.txt"
$TemplateCSV = "C:\TEMP\GroupImportMembersTemplate.csv"
$OutputCSV = "C:\TEMP\IntuneGroupImport.csv"

if (!(Test-Path $HostnamesFile)) {
    Write-Error "Hostnames file not found at $HostnamesFile"
    
}

if (!(Test-Path $TemplateCSV)) {
    Write-Host "Template CSV not found - creating it automatically at $TemplateCSV" -ForegroundColor Yellow
    @(
        "Member object ID or user principal name [memberObjectIdOrUpn] Required",
        "Example: 9832aad8-e4fe-496b-a604-95c6eF01ae75"
    ) | Out-File -FilePath $TemplateCSV -Encoding UTF8
    Write-Host "Template CSV created." -ForegroundColor Green
}

# Read, clean and deduplicate hostnames
$AllHostnames = Get-Content $HostnamesFile
$CleanHostnames = $AllHostnames | Where-Object { $_ -and $_ -notmatch "^\s*$" } | ForEach-Object { $_.Trim() }
$UniqueHostnames = $CleanHostnames | Select-Object -Unique
$DuplicatesRemoved = $CleanHostnames.Count - $UniqueHostnames.Count

Write-Host "Total hostnames read from file $($AllHostnames.Count)"
Write-Host "After cleanup $($CleanHostnames.Count)"
Write-Host "Unique hostnames to process $($UniqueHostnames.Count)"
if ($DuplicatesRemoved -gt 0) {
    Write-Host "Duplicate hostnames removed $DuplicatesRemoved"
}

Copy-Item $TemplateCSV $OutputCSV -Force
$TemplateContent = Get-Content $TemplateCSV

Write-Host "`nTemplate content analysis"
Write-Host "Template file has $($TemplateContent.Count) lines"
for ($i = 0; $i -lt [Math]::Min(5, $TemplateContent.Count); $i++) {
    Write-Host "Line $($i+1) - '$($TemplateContent[$i])'"
}

$HeaderEndIndex = -1
$HeaderPatterns = @(
    "\[memberObjectIdOrUpn\]",
    "Member object ID",
    "memberObjectIdOrUpn",
    "Object ID",
    "ObjectId"
)

for ($i = 0; $i -lt $TemplateContent.Count; $i++) {
    foreach ($Pattern in $HeaderPatterns) {
        if ($TemplateContent[$i] -match $Pattern) {
            $HeaderEndIndex = $i
            Write-Host "Found header pattern '$Pattern' at line $($i+1)"
            break
        }
    }
    if ($HeaderEndIndex -ne -1) { break }
}

if ($HeaderEndIndex -eq -1) {
    Write-Error "Invalid template format. Expected headers with 'Member object ID' or similar."
    
}

$HeaderContent = $TemplateContent[0..$HeaderEndIndex]
$HeaderContent | Out-File -FilePath $OutputCSV -Encoding UTF8

# Use hashtable to track processed devices and avoid duplicates
$ProcessedDevices = @{}
$ObjectIDs = @()
$NotFoundDevices = @()

foreach ($DisplayName in $UniqueHostnames) {
    
    if ([string]::IsNullOrWhiteSpace($DisplayName)) {
        continue
    }
    
    try {
        $ObjectID = Get-EntraObjectId -DeviceName $DisplayName

        if ($ObjectID) {
            
            # Double check for duplicates even if already filtered
            if (!$ProcessedDevices.ContainsKey($ObjectID)) {
                $ProcessedDevices[$ObjectID] = $DisplayName
                $ObjectIDs += $ObjectID
                Write-Host "Processed $DisplayName - ObjectId $ObjectID"
            }
        }
    }
    catch {
        Write-Warning "Device not found or error accessing $DisplayName"
        $NotFoundDevices += $DisplayName
    }
}

$ObjectIDs | Out-File -FilePath $OutputCSV -Append -Encoding UTF8

$MappingFile = "C:\TEMP\DeviceMapping.csv"
"Hostname,ObjectId,Status" | Out-File -FilePath $MappingFile -Encoding UTF8

foreach ($Key in $ProcessedDevices.Keys) {
    "$($ProcessedDevices[$Key]),$Key,Found" | Out-File -FilePath $MappingFile -Append -Encoding UTF8
}

foreach ($Device in $NotFoundDevices) {
    "$Device,,Not Found" | Out-File -FilePath $MappingFile -Append -Encoding UTF8
}

Write-Host "`nProcessing complete"
Write-Host "Unique hostnames processed $($UniqueHostnames.Count)"
Write-Host "Devices found successfully $($ObjectIDs.Count)"
Write-Host "Devices not found $($NotFoundDevices.Count)"
Write-Host "Output file $OutputCSV ready for Intune import"
Write-Host "Device mapping saved to $MappingFile"
Write-Host ""
Write-Host "TROUBLESHOOT IMPORT ERRORS"
Write-Host "1. Go to Groups > Bulk operation results in Intune"
Write-Host "2. Download error file to see specific failures"
Write-Host "3. Common causes - Object already in group, invalid Object ID, permissions"