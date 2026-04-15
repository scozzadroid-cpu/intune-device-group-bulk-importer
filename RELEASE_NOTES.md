# Intune Device Group Bulk Importer - v1.4 (Stable)

## Overview
Intune Device Group Bulk Importer is a PowerShell-based tool for bulk importing devices into Intune groups using Microsoft Graph. This release includes significant improvements to GUI visibility, authentication reliability, and module handling.

## What's New in v1.4

### 🎨 UI/UX Improvements
- **Enhanced Color Visibility**: Improved text contrast in the dark-cool theme for better readability
  - TextBox text: `#E8E8F8` → `#F0F0FF` (brighter)
  - Secondary labels: `#6060A0` → `#9090D0` (more visible)
  - Placeholder text: `#252545` → `#6B6B9B` (readable)
  - Tab items: `#5858A0` → `#8B8BC8` (clearer)

### 🔐 Authentication
- **Interactive Browser as Default**: Changed from unreliable device code flow to interactive browser authentication
- **Fixed Device Code Bug**: Removed `-UseDeviceAuthentication` parameter which had issues in Microsoft Graph v2.26.1
- **WAM Disabled**: Prevents "window handle must be configured" errors
- **Simplified Scope**: Uses only `DeviceManagementManagedDevices.Read.All` (no admin approval required)

### 🔧 Technical Fixes
- **Fixed Module Loading**: All required Microsoft Graph modules now imported correctly:
  - `Microsoft.Graph.Authentication`
  - `Microsoft.Graph.DeviceManagement`
  - `Microsoft.Graph.Identity.DirectoryManagement`
- **Fixed Import-IntuneGroupBySerial.ps1**: Added missing `Import-Module` statements
- **Removed Admin Approval Requirement**: Removed `Device.Read.All` scope that triggered organization admin consent

## Files Included

| File | Purpose |
|------|---------|
| `gui.ps1` | Main WPF GUI application for bulk import operations |
| `gui.bat` | Batch wrapper for easy exe-like execution |
| `Import-IntuneGroupByHostname.ps1` | CLI tool for importing devices by hostname |
| `Import-IntuneGroupBySerial.ps1` | CLI tool for importing devices by serial number |
| `script.ps1` | Reference implementation (legacy) |
| `serial version.ps1` | Serial number processing reference |
| `README.md` | Full documentation |

## Requirements
- Windows 10/11 or Windows Server 2019+
- PowerShell 5.1 or higher
- Microsoft Graph PowerShell modules:
  - Microsoft.Graph.Authentication
  - Microsoft.Graph.DeviceManagement
  - Microsoft.Graph.Identity.DirectoryManagement
- Intune admin permissions
- Azure AD user with appropriate permissions

## Installation

### Quick Start
1. Download all files from this release
2. Extract to a folder (e.g., `C:\Tools\Intune-Bulk-Importer`)
3. Run `gui.ps1` using PowerShell:
   ```powershell
   cd C:\Tools\Intune-Bulk-Importer
   .\gui.ps1
   ```
   Or run the batch wrapper:
   ```cmd
   gui.bat
   ```

### Module Installation
The scripts will automatically attempt to install required modules. If installation fails:
```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force
Install-Module Microsoft.Graph.DeviceManagement -Scope CurrentUser -Force
Install-Module Microsoft.Graph.Identity.DirectoryManagement -Scope CurrentUser -Force
```

## Usage

### GUI Mode (Recommended)
Run `gui.ps1` for the full GUI experience:
1. Click "Connect to Microsoft Graph"
2. Sign in with your admin account (interactive browser popup)
3. Choose import method:
   - **By Hostname**: Paste device hostnames (one per line)
   - **By Serial Number**: Paste serial numbers (one per line)
4. Select output CSV location
5. Click "Run Import"
6. Import the CSV into your Intune device group

### CLI Mode
```powershell
# By hostname
.\Import-IntuneGroupByHostname.ps1

# By serial number
.\Import-IntuneGroupBySerial.ps1
```

## Troubleshooting

### Authentication Issues
- **Error**: "Authentication failed. Please try again."
  - **Solution**: Check your internet connection and admin permissions
  - Ensure your account is not restricted by conditional access

- **Error**: "Window handle must be configured"
  - **Fixed**: WAM is now disabled by default in v1.4

### Module Loading Issues
- **Error**: "The term 'Connect-MgGraph' is not recognized"
  - **Solution**: Modules were not installed correctly
  - Try: `Get-InstalledModule -Name "Microsoft.Graph*" | Uninstall-Module -AllVersions -Force`
  - Then reinstall modules as shown above

### Device Not Found
- Ensure device exists in Intune
- Check device name/serial number spelling
- Verify device has synced with Azure AD

## Known Limitations
- Serial number matching is fuzzy (StartsWith, Contains, Pattern matching)
- Requires DeviceManagementManagedDevices.Read.All permission
- Maximum recommended import size: 5000 devices per operation

## Changelog

### v1.4 (2026-04-15)
- ✅ Enhanced UI colors for better visibility
- ✅ Interactive Browser as default authentication
- ✅ Fixed module loading issues
- ✅ Removed admin approval requirement
- ✅ Fixed device code authentication bug

### v0.9 (Previous)
- Initial release with device code authentication

## Support
For issues or feature requests, please refer to the GitHub repository documentation or the README.md file.

## License
This tool is provided as-is for Intune administration purposes.
