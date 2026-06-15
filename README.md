# Intune Profile Bulk Renamer Tool

[![PowerShell Gallery](https://img.shields.io/powershellgallery/v/Invoke-IntuneProfileManager?logo=powershell&label=PowerShell%20Gallery)](https://www.powershellgallery.com/packages/Invoke-IntuneProfileManager)
[![Downloads](https://img.shields.io/powershellgallery/dt/Invoke-IntuneProfileManager?label=Downloads)](https://www.powershellgallery.com/packages/Invoke-IntuneProfileManager)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

A self-contained **PowerShell 7 + Windows Forms** desktop tool for **bulk renaming and re-describing Microsoft Intune configuration profiles** via the Microsoft Graph API.

Pull every configuration profile from your tenant into an editable grid, change the **display name** and/or **description** — inline, in Excel, or with built-in find & replace — and write the changes back. Only the name and description are ever modified; nothing else about a profile is touched.

> ⚠️ A free tool by **[modernworkspacehub.com](https://modernworkspacehub.com)**. Provided **"as is", without warranty of any kind** — it modifies live data in your Intune tenant, so use at your own risk. See [Disclaimer](#disclaimer).

---

## Table of contents

- [Screenshots](#screenshots)
- [What it does](#what-it-does)
- [What it does *not* do](#what-it-does-not-do)
- [Profile types covered](#profile-types-covered)
- [Requirements](#requirements)
- [Permissions](#permissions)
- [Install](#install)
- [Usage](#usage)
- [CSV format](#csv-format)
- [Find & Replace](#find--replace)
- [Backups & restore](#backups--restore)
- [Logs](#logs)
- [Safety notes](#safety-notes)
- [Troubleshooting](#troubleshooting)
- [Disclaimer](#disclaimer)

---

## Screenshots

<!--
  Add a few screenshots here. Drop the image files into a `docs/` folder in the repo
  and update the paths below. Keep it to two or three — the main window and a key feature
  are usually enough. Suggested shots: the main window with profiles loaded, and Find & Replace.
-->

![Main window — profiles loaded in the editable grid](docs/screenshot-main.png)

![Find & Replace dialog](docs/screenshot-find-replace.png)

<!-- Optional third shot, e.g. a dry-run / apply result in the activity log:
![Dry-run preview in the activity log](docs/screenshot-dryrun.png)
-->

---

## What it does

- **Connects to Microsoft Graph** interactively using the Microsoft Graph PowerShell SDK.
- **Pulls all configuration profiles** across four families (see below) into an editable data grid.
- **Shows the specific template kind** in the Type column, e.g. `Device Configuration (VPN)`, `(SCEP certificate)`, `(Domain join)`, `(Wi-Fi)`.
- **Edits inline** — change *New Name* / *New Description* directly in the grid; changed cells are highlighted.
- **Exports to CSV** for bulk editing in Excel, then **imports the CSV back** in.
- **Find & Replace** across names/descriptions — literal or regex, with a one-click "strip trailing version" preset (e.g. removing ` v3.9`, `-2.8`).
- **Dry-run mode** previews exactly what would change without calling Graph.
- **JSON backup & restore** — snapshot current names/descriptions and revert if needed (an automatic safety backup is also taken before every apply).
- **Applies changes** by sending a minimal `PATCH` to the correct Graph endpoint for **only the rows that changed**.
- **Logs everything** to a real-time activity log and a daily log file in the repo.

## What it does *not* do

- ❌ Does **not** change anything other than **display name** and **description**. Settings, assignments, scope tags, platform, etc. are never modified.
- ❌ Does **not** create, delete, duplicate, or assign profiles.
- ❌ Does **not** touch profile *settings/payloads* — it is a metadata (name/description) editor only.
- ❌ Does **not** manage compliance policies, app protection/configuration policies, scripts, remediations, autopilot profiles, or enrollment configurations.
- ❌ Does **not** support app-only/unattended authentication — sign-in is interactive (delegated).
- ❌ Is **not** a Microsoft product and is **not** supported by Microsoft.

## Profile types covered

All endpoints use Microsoft Graph **beta**, which surfaces every derived template type (the `v1.0` list omits many, such as certificate and health-monitoring profiles).

| Type shown in tool | Graph collection | Notes |
|---|---|---|
| **Settings Catalog** | `deviceManagement/configurationPolicies` | Uses `name` / `description` |
| **Device Configuration** | `deviceManagement/deviceConfigurations` | All templates: device restrictions, domain join, Wi-Fi, VPN, SCEP/PKCS/trusted certificates, health monitoring, kiosk, email, custom OMA-URI, etc. |
| **Administrative Template** | `deviceManagement/groupPolicyConfigurations` | Imported/ADMX-backed templates |
| **Template / Baseline** | `deviceManagement/intents` | Endpoint security & security baselines |

## Requirements

- **Windows** (the UI is built on Windows Forms).
- **PowerShell 7+** (`pwsh.exe`). It will refuse to run on Windows PowerShell 5.1.
- **Microsoft Graph PowerShell SDK** — only the authentication module is required:
  ```powershell
  Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
  ```
- An Intune-licensed Entra ID tenant and an account with sufficient permissions (below).
- No other third-party dependencies.

## Permissions

- **Delegated Graph scope:** `DeviceManagementConfiguration.ReadWrite.All`
  - The tool requests this scope at sign-in. The first time, an admin (or the user, if user consent is allowed) must consent.
- **Directory role:** an account that can read and modify Intune device configuration — typically **Intune Administrator** (or a custom role with the equivalent read/update rights). Global Administrator works but is not required.
- If your role can only see a subset of profile types, the others are skipped with a warning in the log rather than failing the whole pull.

## Install

### Option A — PowerShell Gallery (recommended)

Published as a script on the [PowerShell Gallery](https://www.powershellgallery.com/packages/Invoke-IntuneProfileManager):

```powershell
# Install the prerequisite Graph module (one-off)
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser

# Install the tool
Install-Script -Name Invoke-IntuneProfileManager -Scope CurrentUser

# Run it (from PowerShell 7)
Invoke-IntuneProfileManager.ps1
```

Update later with `Update-Script -Name Invoke-IntuneProfileManager`.

> `Install-Script` saves to your user scripts folder (which is on `PATH`), so you can launch it by name. If the name isn't found, run it with its full path or `& "$(Split-Path (Get-InstalledScript Invoke-IntuneProfileManager).InstalledLocation)\Invoke-IntuneProfileManager.ps1"`.

### Option B — Download / clone

1. Install the prerequisite module: `Install-Module Microsoft.Graph.Authentication -Scope CurrentUser`.
2. Download or clone this repository.
3. Run the script from **PowerShell 7**:

   ```powershell
   pwsh .\Invoke-IntuneProfileManager.ps1
   ```

   > Tip: launch from a fresh `pwsh` window for best results.

The script is a single self-contained `.ps1` file. It creates `Logs\` and `Backups\` folders next to itself on first run.

## Usage

1. **Connect** — click **Connect to Intune** and sign in. The status bar shows the connected account and tenant.
2. **Pull** — click **Pull** to load all configuration profiles into the grid.
3. **Edit** — change values in the **New Name** / **New Description** columns. Changed cells are highlighted amber. You can:
   - Edit directly in the grid, **or**
   - **Export** to CSV, edit in Excel, then **Import** the CSV back, **or**
   - Use **Find/Replace** for bulk changes (see below).
4. **(Optional) Backup** — click **Backup** to save a JSON snapshot you can restore later.
5. **(Optional) Dry run** — tick **Dry run** to preview what would change without writing anything.
6. **Apply** — click **Apply**. You'll be asked to confirm. Only rows where the name or description differs from the current value are sent to Graph; each result is logged. (A safety backup is written automatically before any real apply.)

> Toolbar buttons have tooltips describing each action.

## CSV format

Export/import uses these columns:

| Column | Editable | Purpose |
|---|---|---|
| `ProfileId` | No | Graph object ID (do not change) |
| `ProfileType` | No | Profile family/template kind (used to route the update) |
| `CurrentName` | No | Current display name (for reference/diffing) |
| `NewName` | **Yes** | The new display name to apply |
| `CurrentDescription` | No | Current description (for reference/diffing) |
| `NewDescription` | **Yes** | The new description to apply |

Edit only the **NewName** and **NewDescription** columns. On import, rows are validated; malformed or unrecognised rows are skipped and reported in the log. A change is applied only where a New value differs from its Current value.

## Find & Replace

Operates on the **New Name** / **New Description** columns in the grid (nothing is sent to Intune until you Apply):

- **Literal** or **regular expression** matching, with an optional case-sensitive toggle.
- Target **New Name**, **New Description**, or both.
- Scope to **all rows** or **selected rows only**.
- Leave **Replace with** blank to **delete** the matched text.
- **Preview** counts matches before committing.
- **Strip trailing version** preset removes a trailing version postfix such as ` v3.9`, ` 3.1`, `-2.8`, `V10.0.1` (regex: `[\s_\-]*[vV]?\d+(\.\d+)+\s*$`). Tweak it or write your own afterwards.

## Backups & restore

- **Backup** writes a timestamped JSON snapshot of the current names and descriptions to `Backups\`.
- An **automatic** snapshot (`*_auto-preapply.json`) is written before every non-dry-run apply.
- **Restore** reads a backup and writes the saved names/descriptions back to Intune (respects Dry run, with confirmation and per-profile logging).

## Logs

- A live **Activity Log** is shown in the app.
- Everything is also appended to a daily file: `Logs\IntuneProfileManager_yyyyMMdd.log`.

## Safety notes

- **Always run a Dry run first** on a large batch, and/or take a **Backup**, before applying.
- Renames change how profiles appear in the Intune portal and any reporting that keys off display name — coordinate with your team before bulk renaming.
- The tool never blanks a name: rows with an empty **New Name** are skipped.
- Test against a small selection (or a non-production tenant) before running tenant-wide.

## Troubleshooting

| Symptom | Likely cause / fix |
|---|---|
| "This tool requires PowerShell 7" | You launched it in Windows PowerShell 5.1. Use `pwsh.exe`. |
| "Microsoft Graph SDK not found" | Run `Install-Module Microsoft.Graph.Authentication -Scope CurrentUser`. |
| A profile type is missing after Pull | Your role may lack access to that type — check the log for a skipped/permission warning. |
| `403 / Forbidden` on apply | Missing consent or insufficient role for `DeviceManagementConfiguration.ReadWrite.All`. |
| `404 / NotFound` on apply | The profile was deleted since the last pull — Pull again. |

## Disclaimer

This tool is provided **"as is", without warranty of any kind**, express or implied, including but not limited to the warranties of merchantability, fitness for a particular purpose, and non-infringement. In no event shall the author or modernworkspacehub.com be liable for any claim, damages, or other liability arising from the use of this tool.

**It modifies live data in your Microsoft Intune tenant. Use at your own risk.** You are responsible for testing it and for any changes it makes. This project is not affiliated with, endorsed by, or supported by Microsoft.

---

*Made by [modernworkspacehub.com](https://modernworkspacehub.com). Part of the same toolset as Win32Forge.*
