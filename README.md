# Intune Bulk Group Import — PowerShell Tools

Two PowerShell scripts and a WPF GUI for bulk-importing devices into **Azure AD / Entra ID groups** using the Microsoft Graph API.  
Each tool reads a list of devices, resolves their Object IDs, and produces a CSV ready for the **Intune Group Bulk Import** wizard.

---

## GUI — `Import-IntuneGroupGUI.exe` / `Import-IntuneGroupGUI.ps1`

A dark-themed WPF desktop application that wraps both workflows in a point-and-click interface.  
All Graph calls run in a background runspace — the UI never freezes.

### How to run

**Option A — Executable (no PowerShell setup required)**

1. Download `Import-IntuneGroupGUI.exe`.
2. Double-click it. Windows may show a SmartScreen warning the first time — click **More info → Run anyway**.

**Option B — PowerShell script**

```powershell
powershell.exe -ExecutionPolicy Bypass -STA -File Import-IntuneGroupGUI.ps1
```

> The script self-restarts in STA mode if needed, so you can also right-click → *Run with PowerShell*.

---

### Step-by-step usage

#### 1. Prerequisites check

On first launch an overlay checks for the required modules:

| Module | Used for |
|---|---|
| `Microsoft.Graph.Authentication` | `Connect-MgGraph` |
| `Microsoft.Graph.DeviceManagement` | Hostname → Object ID (Intune) |
| `Microsoft.Graph.Identity.DirectoryManagement` | Serial → Object ID (Entra ID) |

Missing modules show a **red dot**. Click **Install** — the app installs them automatically (requires internet; may prompt for the NuGet provider on first run).  
When all dots turn green click **Continue →**.

#### 2. Connect to Microsoft Graph

Select the authentication method, then click **Connect to Microsoft Graph**:

| Method | When to use |
|---|---|
| **Device code** *(default)* | Headless/terminal sessions, MFA with Authenticator app, no browser on the same machine |
| **Interactive (browser popup)** | Standard workstation with a browser available |

**Device code flow:** the URL and one-time code appear in the Log panel:

```
To sign in, use a web browser to open https://microsoft.com/devicelogin
and enter the code XXXXXXXX to authenticate.
```

Open the URL in any browser, enter the code, and sign in with your Azure AD account.

**Scopes requested:**

- `DeviceManagementManagedDevices.Read.All` — used by the Hostname tab
- `Device.Read.All` — used by the Serial tab

After sign-in the auth bar shows **Connected** in green with the signed-in account and tenant ID.

#### 3. Hostname tab

Resolves device names against the **Intune `managedDevices`** endpoint.

1. Paste hostnames directly into the text area (one per line), **or** click **Browse…** to pick a `.txt` file and then **Load**.
2. Set the output path in the *Output* field (default `C:\TEMP\IntuneGroupImport.csv`), or click **…** to browse.
3. Click **Run Import**.

The app downloads all Intune managed devices once, then matches each hostname.  
Progress is shown in the bar and log. When complete a summary popup appears and two files are written:

| File | Contents |
|---|---|
| `IntuneGroupImport.csv` | Ready for Azure Portal bulk import |
| `IntuneGroupImport_mapping.csv` | Hostname ↔ Object ID audit log |

Log output per device: `FOUND    DESKTOP-ABC  ->  <ObjectID>` or `NOT FOUND  DESKTOP-XYZ`.

#### 4. Serial Number tab

Resolves serial numbers against the **Entra ID `devices`** endpoint using four fallback strategies per serial:

| Strategy | Logic |
|---|---|
| **Exact** | `displayName` exactly equals the serial |
| **StartsWith** | `displayName` starts with the serial |
| **Contains** | `displayName` contains the serial |
| **Fuzzy** | Case-insensitive regex match |

1. Paste serial numbers (one per line), or **Browse…** + **Load** from a `.txt` file.
2. Set the output path, click **Run Import**.

Log output: `[Exact]  C02XG2JH  ->  MacBook-Pro  ->  <ObjectID>` or `NOT FOUND  BADSERIAL`.  
Duplicate Object IDs are removed automatically.

#### 5. Upload to Azure Portal

1. Open **Azure Portal → Groups → [your group] → Members → Bulk operations → Import members**.
2. Upload the output CSV file.
3. Check **Bulk operation results** for any errors (Object ID not found, already a member, etc.).

---

## Scripts (CLI)

### `Import-IntuneGroupByHostname.ps1`

Resolves devices by **hostname** using the Intune `managedDevices` endpoint.

**Usage:**

```powershell
# 1. Create C:\TEMP\pc.txt with one hostname per line
# 2. Run:
powershell.exe -ExecutionPolicy Bypass -File Import-IntuneGroupByHostname.ps1
```

| | |
|---|---|
| **Scope** | `DeviceManagementManagedDevices.Read.All` |
| **Input** | `C:\TEMP\pc.txt` |
| **Output** | `C:\TEMP\IntuneGroupImport.csv` + `C:\TEMP\DeviceMapping.csv` |

---

### `Import-IntuneGroupBySerial.ps1`

Resolves devices by **serial number** using the Entra ID `devices` endpoint.

**Usage:**

```powershell
# 1. Create C:\TEMP\serials.txt with one serial per line
# 2. Run:
powershell.exe -ExecutionPolicy Bypass -File Import-IntuneGroupBySerial.ps1
```

| | |
|---|---|
| **Scope** | `Device.Read.All` |
| **Input** | `C:\TEMP\serials.txt` |
| **Output** | `C:\TEMP\IntuneGroupImport.csv` |

The script tries the same four fallback strategies as the GUI (Exact → StartsWith → Contains → Fuzzy) and prints a per-strategy summary at the end.

---

## Requirements

- Windows 10 / 11
- PowerShell 5.1 or PowerShell 7+
- Microsoft.Graph PowerShell SDK (the GUI installs modules automatically on first run)

Manual install:

```powershell
Install-Module Microsoft.Graph.Authentication -Force -Scope CurrentUser
Install-Module Microsoft.Graph.DeviceManagement -Force -Scope CurrentUser
Install-Module Microsoft.Graph.Identity.DirectoryManagement -Force -Scope CurrentUser
```

---

## Notes

- All tools use `Invoke-MgGraphRequest` with manual `@odata.nextLink` pagination instead of the `-All` SDK flag, which throws `AggregateException` on some SDK versions.
- `MSAL_DISABLE_WAM=1` is set automatically at startup to prevent the WAM window-handle error on Windows 11.
- The hostname lookup returns the **Intune managed device ID** (from `managedDevices`). This is sufficient for Intune-targeted groups. If you need the true Entra Object ID, grant `Device.Read.All` and uncomment Option A inside `Get-EntraObjectId` in `script.ps1`.
- Duplicate Object IDs are removed before writing the output CSV in all tools.

---

## Author

Lorenzo Scozzafava
