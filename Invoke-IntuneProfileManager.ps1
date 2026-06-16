#Requires -Version 7.0

<#PSScriptInfo

.VERSION 1.1.0

.GUID 041b1471-ad40-45b7-9fb0-81a12f91cd19

.AUTHOR Alex Durrant

.COMPANYNAME modernworkspacehub.com

.COPYRIGHT (c) 2026 Alex Durrant / modernworkspacehub.com. All rights reserved.

.TAGS Intune Graph MicrosoftGraph DeviceManagement ConfigurationProfiles BulkRename Rename MEM Windows

.LICENSEURI https://github.com/durrante/Intune-Profile-Bulk-Renamer-Tool/blob/main/LICENSE

.PROJECTURI https://github.com/durrante/Intune-Profile-Bulk-Renamer-Tool

.ICONURI

.EXTERNALMODULEDEPENDENCIES

.REQUIREDSCRIPTS

.EXTERNALSCRIPTDEPENDENCIES

.RELEASENOTES
1.1.0 - Expanded from configuration profiles to ~20 Intune content types (compliance policies,
app protection & configuration, PowerShell/remediation/macOS scripts, assignment filters,
Autopilot profiles, device categories, driver/feature/quality updates, and opt-in Entra groups).
Added a Content Types picker to choose what each Pull fetches, expanded Graph scopes, and fixed
content-type selection persistence. 1.0.0 - Initial release. Throughout, only display name and
description are ever modified; nothing else about an item is changed.

.PRIVATEDATA

#>

<#
.SYNOPSIS
    Intune Profile Manager — bulk rename & re-describe Intune configuration profiles.

.DESCRIPTION
    A self-contained Windows Forms (PowerShell 7+) desktop tool that connects to Microsoft
    Graph and bulk-edits the display name and description of Intune objects. Only the name
    and description are ever changed - nothing else about an item is touched.

    Covers ~20 content types (all via Graph beta, so every derived template is surfaced):
      - Settings Catalog and Device Configuration (every template: device restrictions, domain
        join, Wi-Fi, VPN, SCEP/PKCS/trusted certificates, health monitoring, kiosk, custom, ...)
      - Administrative Templates and Templates / security baselines (intents)
      - Compliance policies
      - App protection (iOS/Android) and app configuration (managed apps/devices)
      - PowerShell, remediation and macOS shell scripts
      - Assignment filters, Autopilot profiles, device categories
      - Driver / feature / quality update profiles and quality update policies
      - Entra ID groups (opt-in)

    Workflow:
      1. Connect to Microsoft Graph (delegated sign-in; consents the scopes for the above).
      2. Pick which content types to pull, then Pull them into an editable grid.
      3. Edit New Name / New Description inline, use Find & Replace, or Export to CSV,
         edit in Excel, and Import back.
      4. Apply - only items whose name or description changed are PATCHed.
      5. Dry-run mode previews changes without calling Graph. JSON backup & restore included.

    Project, documentation and issues:
        https://github.com/durrante/Intune-Profile-Bulk-Renamer-Tool

    A modernworkspacehub.com tool, part of the same toolset as Win32Forge
    (https://github.com/durrante/Win32Forge).

    Provided "as is", without warranty of any kind - use at your own risk. Not affiliated
    with, endorsed by, or supported by Microsoft.

    Requires the Microsoft Graph PowerShell SDK authentication module:
        Install-Module Microsoft.Graph.Authentication -Scope CurrentUser

.EXAMPLE
    pwsh .\Invoke-IntuneProfileManager.ps1
#>

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

# Belt-and-braces check in case the #Requires line is somehow bypassed
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Error ("This tool requires PowerShell 7 (pwsh.exe).`n" +
                 "You are running PowerShell $($PSVersionTable.PSVersion).`n`n" +
                 "Start the tool with:  pwsh `"$PSCommandPath`"")
    exit 1
}

# ─────────────────────────────────────────────────────────────────────────────
#region Assemblies & palette
# ─────────────────────────────────────────────────────────────────────────────
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()
# Can only be set once per process — throws if a WinForms object already exists
# (e.g. when re-running in the same PowerShell session). Safe to ignore in that case.
try { [System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false) } catch {}

function C { param([string]$Hex) [System.Drawing.ColorTranslator]::FromHtml($Hex) }

# Palette lifted directly from Win32Forge
$Theme = @{
    GradLeft     = C '#0693E3'   # header gradient start (blue)
    GradRight    = C '#9B51E0'   # header gradient end   (purple)
    SubtleText   = C '#D4C5F9'   # header subtitle
    StatusBg     = C '#F0F0F0'
    Border       = C '#DDDDDD'
    ToolbarBg    = C '#F8F8F8'
    ToolbarLine  = C '#E0E0E0'
    FooterBg     = C '#F5F5F5'
    FooterText   = C '#666666'
    FooterFaint  = C '#AAAAAA'
    Primary      = C '#4A2B8F'   # main purple
    PrimaryDark  = C '#2D1B69'
    AccentBlue   = C '#5BA3E8'
    DeepPurple   = C '#3A2673'
    DotRed       = C '#D32F2F'
    DotGreen     = C '#2E7D32'
    GridHeaderBg = C '#F0EBF9'
    GridHeaderFg = C '#4A2B8F'
    GridAltRow   = C '#FAFAFA'
    GridSelBg    = C '#E8DEFF'
    GridSelFg    = C '#2D1B69'
    GridLine     = C '#DDDDDD'
    ChangedBg    = C '#FFF4D6'   # amber tint for pending-change cells
    ChangedFg    = C '#8A5A00'
    LogOk        = C '#2E7D32'
    LogWarn      = C '#B7791F'
    LogFail      = C '#C62828'
    LogInfo      = C '#333333'
    White        = [System.Drawing.Color]::White
    BtnText      = C '#333333'
}

$FontUI      = New-Object System.Drawing.Font('Segoe UI', 9)
$FontUIBold  = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$FontTitle   = New-Object System.Drawing.Font('Segoe UI Light', 18, [System.Drawing.FontStyle]::Regular)
$FontSub     = New-Object System.Drawing.Font('Segoe UI', 8.5)
$FontMono    = New-Object System.Drawing.Font('Consolas', 9.5)

# Graph endpoints — beta is used throughout because it surfaces every derived type.
$GraphBeta = 'https://graph.microsoft.com/beta'
$GraphV1   = 'https://graph.microsoft.com/v1.0'

# Delegated scopes requested at sign-in. Together they cover every content type below.
# If the signed-in admin hasn't consented to one, the affected types are simply skipped
# on pull (logged as a warning) rather than failing the whole operation.
$RequiredScopes = @(
    'DeviceManagementConfiguration.ReadWrite.All'   # config, compliance, admin templates, baselines, filters, update profiles
    'DeviceManagementScripts.ReadWrite.All'         # PowerShell / remediation / macOS shell scripts
    'DeviceManagementApps.ReadWrite.All'            # app protection & app configuration policies
    'DeviceManagementServiceConfig.ReadWrite.All'   # Autopilot profiles, enrollment configurations
    'Group.ReadWrite.All'                           # Entra ID groups (opt-in content type)
)

# Content type catalogue — every Intune object family the tool can rename / re-describe.
# Keyed by the friendly name shown in the grid, CSV and Content Types picker.
#   Base/Collection : Graph endpoint
#   Select          : $select clause (keeps payloads small)
#   NameProp        : property holding the display name when reading
#   PatchNameProp   : property name the new display name is sent as on PATCH
#   NeedsODataType  : polymorphic collections need the derived @odata.type on PATCH
#   Subtype         : show the specific template kind in brackets (Device Configuration only)
#   Default         : pre-checked in the Content Types picker (Entra groups are opt-in)
$ContentTypes = [ordered]@{
    'Settings Catalog'            = @{ Base=$GraphBeta; Collection='deviceManagement/configurationPolicies';          Select='id,name,description';        NameProp='name';        PatchNameProp='name';        NeedsODataType=$false; Subtype=$false; Default=$true }
    'Device Configuration'        = @{ Base=$GraphBeta; Collection='deviceManagement/deviceConfigurations';           Select='id,displayName,description'; NameProp='displayName'; PatchNameProp='displayName'; NeedsODataType=$true;  Subtype=$true;  Default=$true }
    'Administrative Template'     = @{ Base=$GraphBeta; Collection='deviceManagement/groupPolicyConfigurations';      Select='id,displayName,description'; NameProp='displayName'; PatchNameProp='displayName'; NeedsODataType=$false; Subtype=$false; Default=$true }
    'Template / Baseline'         = @{ Base=$GraphBeta; Collection='deviceManagement/intents';                        Select='id,displayName,description'; NameProp='displayName'; PatchNameProp='displayName'; NeedsODataType=$false; Subtype=$false; Default=$true }
    'Compliance Policy'           = @{ Base=$GraphBeta; Collection='deviceManagement/deviceCompliancePolicies';       Select='id,displayName,description'; NameProp='displayName'; PatchNameProp='displayName'; NeedsODataType=$true;  Subtype=$false; Default=$true }
    'PowerShell Script'           = @{ Base=$GraphBeta; Collection='deviceManagement/deviceManagementScripts';        Select='id,displayName,description'; NameProp='displayName'; PatchNameProp='displayName'; NeedsODataType=$false; Subtype=$false; Default=$true }
    'Remediation Script'          = @{ Base=$GraphBeta; Collection='deviceManagement/deviceHealthScripts';            Select='id,displayName,description'; NameProp='displayName'; PatchNameProp='displayName'; NeedsODataType=$false; Subtype=$false; Default=$true }
    'macOS Shell Script'          = @{ Base=$GraphBeta; Collection='deviceManagement/deviceShellScripts';             Select='id,displayName,description'; NameProp='displayName'; PatchNameProp='displayName'; NeedsODataType=$false; Subtype=$false; Default=$true }
    'App Protection (iOS)'        = @{ Base=$GraphBeta; Collection='deviceAppManagement/iosManagedAppProtections';     Select='id,displayName,description'; NameProp='displayName'; PatchNameProp='displayName'; NeedsODataType=$false; Subtype=$false; Default=$true }
    'App Protection (Android)'    = @{ Base=$GraphBeta; Collection='deviceAppManagement/androidManagedAppProtections'; Select='id,displayName,description'; NameProp='displayName'; PatchNameProp='displayName'; NeedsODataType=$false; Subtype=$false; Default=$true }
    'App Config (managed apps)'   = @{ Base=$GraphBeta; Collection='deviceAppManagement/targetedManagedAppConfigurations'; Select='id,displayName,description'; NameProp='displayName'; PatchNameProp='displayName'; NeedsODataType=$false; Subtype=$false; Default=$true }
    'App Config (managed devices)'= @{ Base=$GraphBeta; Collection='deviceAppManagement/mobileAppConfigurations';     Select='id,displayName,description'; NameProp='displayName'; PatchNameProp='displayName'; NeedsODataType=$true;  Subtype=$false; Default=$true }
    'Assignment Filter'           = @{ Base=$GraphBeta; Collection='deviceManagement/assignmentFilters';              Select='id,displayName,description'; NameProp='displayName'; PatchNameProp='displayName'; NeedsODataType=$false; Subtype=$false; Default=$true }
    'Autopilot Profile'           = @{ Base=$GraphBeta; Collection='deviceManagement/windowsAutopilotDeploymentProfiles'; Select='id,displayName,description'; NameProp='displayName'; PatchNameProp='displayName'; NeedsODataType=$true; Subtype=$false; Default=$true }
    'Device Category'             = @{ Base=$GraphBeta; Collection='deviceManagement/deviceCategories';               Select='id,displayName,description'; NameProp='displayName'; PatchNameProp='displayName'; NeedsODataType=$false; Subtype=$false; Default=$true }
    'Driver Update Profile'       = @{ Base=$GraphBeta; Collection='deviceManagement/windowsDriverUpdateProfiles';    Select='id,displayName,description'; NameProp='displayName'; PatchNameProp='displayName'; NeedsODataType=$false; Subtype=$false; Default=$true }
    'Feature Update Profile'      = @{ Base=$GraphBeta; Collection='deviceManagement/windowsFeatureUpdateProfiles';   Select='id,displayName,description'; NameProp='displayName'; PatchNameProp='displayName'; NeedsODataType=$false; Subtype=$false; Default=$true }
    'Quality Update Profile'      = @{ Base=$GraphBeta; Collection='deviceManagement/windowsQualityUpdateProfiles';   Select='id,displayName,description'; NameProp='displayName'; PatchNameProp='displayName'; NeedsODataType=$false; Subtype=$false; Default=$true }
    'Quality Update Policy'       = @{ Base=$GraphBeta; Collection='deviceManagement/windowsQualityUpdatePolicies';   Select='id,displayName,description'; NameProp='displayName'; PatchNameProp='displayName'; NeedsODataType=$false; Subtype=$false; Default=$true }
    'Entra Group'                 = @{ Base=$GraphV1;   Collection='groups';                                          Select='id,displayName,description'; NameProp='displayName'; PatchNameProp='displayName'; NeedsODataType=$false; Subtype=$false; Default=$false }
}

# ── Paths (logs + backups live in the same repo as the script) ──────────────
$script:ScriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$script:LogDir    = Join-Path $script:ScriptDir 'Logs'
$script:BackupDir = Join-Path $script:ScriptDir 'Backups'
foreach ($d in @($script:LogDir, $script:BackupDir)) {
    if (-not (Test-Path $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
}
$script:LogFile = Join-Path $script:LogDir "IntuneProfileManager_$(Get-Date -Format 'yyyyMMdd').log"

# ── Shared state ────────────────────────────────────────────────────────────
$script:Connected = $false
$script:Loading   = $false          # suppresses highlight churn during bulk grid load
$script:OdataMap  = @{}             # ProfileId -> @odata.type (polymorphic types only)

# Which content types the next Pull will fetch (chosen in the Content Types picker).
# Defaults to every type flagged Default=$true (i.e. everything except Entra groups).
# NOTE: global scope — the picker saves from inside a closure, and a $script: write
# inside a closure lands in the closure's own scope, not the real script scope.
$global:IPMSelectedContentTypes = [System.Collections.Generic.List[string]]::new()
foreach ($k in $ContentTypes.Keys) { if ($ContentTypes[$k].Default) { $global:IPMSelectedContentTypes.Add($k) } }

#endregion

# ─────────────────────────────────────────────────────────────────────────────
#region Button styling helpers (match Win32Forge tile / tool button look)
# ─────────────────────────────────────────────────────────────────────────────
# Darken (factor < 1) or lighten (factor > 1) a colour — used for solid hover states
function Get-Shade {
    param([System.Drawing.Color]$Color, [double]$Factor)
    $r = [Math]::Max(0, [Math]::Min(255, [int]($Color.R * $Factor)))
    $g = [Math]::Max(0, [Math]::Min(255, [int]($Color.G * $Factor)))
    $b = [Math]::Max(0, [Math]::Min(255, [int]($Color.B * $Factor)))
    return [System.Drawing.Color]::FromArgb($r, $g, $b)
}

# Consistent sizing shared by every toolbar / status button so the row lines up cleanly.
# -Compact gives a shorter, tighter button for the slim status bar.
function Set-ButtonMetrics {
    param([System.Windows.Forms.Button]$Btn, [switch]$Compact)
    $Btn.FlatStyle    = 'Flat'
    $Btn.Font         = $FontUI
    $Btn.Cursor       = 'Hand'
    $Btn.AutoSize     = $true
    $Btn.AutoSizeMode = 'GrowAndShrink'
    $Btn.TextAlign    = 'MiddleCenter'
    if ($Compact) {
        $Btn.Padding     = New-Object System.Windows.Forms.Padding(14, 3, 14, 3)
        $Btn.MinimumSize = New-Object System.Drawing.Size(0, 26)
        $Btn.Margin      = New-Object System.Windows.Forms.Padding(0, 0, 0, 0)
    } else {
        $Btn.Padding     = New-Object System.Windows.Forms.Padding(15, 5, 15, 5)
        $Btn.MinimumSize = New-Object System.Drawing.Size(0, 30)
        $Btn.Margin      = New-Object System.Windows.Forms.Padding(0, 12, 8, 0)
    }
}

function Set-PrimaryButton {
    param([System.Windows.Forms.Button]$Btn, [System.Drawing.Color]$Back, [switch]$Compact)
    Set-ButtonMetrics -Btn $Btn -Compact:$Compact
    $Btn.BackColor = $Back
    $Btn.ForeColor = $Theme.White
    $Btn.FlatAppearance.BorderSize = 0
    $Btn.FlatAppearance.MouseOverBackColor = Get-Shade -Color $Back -Factor 0.88
    $Btn.FlatAppearance.MouseDownBackColor = Get-Shade -Color $Back -Factor 0.75
}

function Set-ToolButton {
    param([System.Windows.Forms.Button]$Btn, [switch]$Compact)
    Set-ButtonMetrics -Btn $Btn -Compact:$Compact
    $Btn.BackColor = $Theme.White
    $Btn.ForeColor = $Theme.BtnText
    $Btn.FlatAppearance.BorderColor = C '#CCCCCC'
    $Btn.FlatAppearance.BorderSize  = 1
    $Btn.FlatAppearance.MouseOverBackColor = C '#ECECEC'
    $Btn.FlatAppearance.MouseDownBackColor = C '#DDDDDD'
}

# Thin vertical divider between logical button groups in a FlowLayoutPanel
function Add-FlowSeparator {
    param([System.Windows.Forms.FlowLayoutPanel]$Flow)
    $sep = New-Object System.Windows.Forms.Panel
    $sep.Size      = New-Object System.Drawing.Size(1, 24)
    $sep.BackColor = C '#C4C4C4'
    $sep.Margin    = New-Object System.Windows.Forms.Padding(6, 15, 13, 0)
    $Flow.Controls.Add($sep)
    return $sep
}

# Build the application icon at runtime from the brand gradient (no external .ico needed):
# a rounded square in blue→purple with three white "rename" bars. Used for the title bar
# and taskbar so the tool presents like a real app.
function New-AppIcon {
    $size = 32
    $bmp  = New-Object System.Drawing.Bitmap($size, $size)
    $g    = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = 'AntiAlias'

    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $d = 14   # corner diameter
    $path.AddArc(0, 0, $d, $d, 180, 90)
    $path.AddArc($size - $d, 0, $d, $d, 270, 90)
    $path.AddArc($size - $d, $size - $d, $d, $d, 0, 90)
    $path.AddArc(0, $size - $d, $d, $d, 90, 90)
    $path.CloseFigure()

    $rect  = New-Object System.Drawing.Rectangle(0, 0, $size, $size)
    $brush = New-Object System.Drawing.Drawing2D.LinearGradientBrush($rect, $Theme.GradLeft, $Theme.GradRight, 0.0)
    $g.FillPath($brush, $path)

    $white = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
    foreach ($y in 8, 15, 22) { $g.FillRectangle($white, 7, $y, 18, 3) }

    $brush.Dispose(); $white.Dispose(); $path.Dispose(); $g.Dispose()
    $hicon = $bmp.GetHicon()
    return [System.Drawing.Icon]::FromHandle($hicon)
}
#endregion

# ─────────────────────────────────────────────────────────────────────────────
#region Form scaffold
# ─────────────────────────────────────────────────────────────────────────────
$form = New-Object System.Windows.Forms.Form
$form.Text          = 'Intune Profile Manager'
$form.Size          = New-Object System.Drawing.Size(1240, 760)
$form.MinimumSize   = New-Object System.Drawing.Size(1080, 560)
$form.StartPosition = 'CenterScreen'
$form.BackColor     = $Theme.White
$form.Font          = $FontUI
try { $form.Icon = New-AppIcon } catch {}

# Root layout — 5 stacked rows
$root = New-Object System.Windows.Forms.TableLayoutPanel
$root.Dock        = 'Fill'
$root.ColumnCount = 1
$root.RowCount    = 5
$root.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 64)))  | Out-Null # header
$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 42)))  | Out-Null # status
$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 54)))  | Out-Null # toolbar
$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))  | Out-Null # split
$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 28)))  | Out-Null # footer
$form.Controls.Add($root)

# ── HEADER (gradient, painted) ───────────────────────────────────────────────
$header = New-Object System.Windows.Forms.Panel
$header.Dock = 'Fill'
$header.Add_Paint({
    param($s, $e)
    $rect  = $s.ClientRectangle
    if ($rect.Width -le 0 -or $rect.Height -le 0) { return }
    $brush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
        $rect, $Theme.GradLeft, $Theme.GradRight, 0.0)
    $e.Graphics.FillRectangle($brush, $rect)
    $brush.Dispose()
    $e.Graphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit
    $titleBrush = New-Object System.Drawing.SolidBrush($Theme.White)
    $subBrush   = New-Object System.Drawing.SolidBrush($Theme.SubtleText)
    $e.Graphics.DrawString('Intune Profile Manager', $FontTitle, $titleBrush, 18, 8)
    $e.Graphics.DrawString('Bulk rename & re-describe Intune configuration profiles  •  modernworkspacehub.com',
        $FontSub, $subBrush, 21, 40)
    $titleBrush.Dispose()
    $subBrush.Dispose()
})
$root.Controls.Add($header, 0, 0)

# ── STATUS BAR ───────────────────────────────────────────────────────────────
$statusBar = New-Object System.Windows.Forms.Panel
$statusBar.Dock      = 'Fill'
$statusBar.BackColor = $Theme.StatusBg
$statusBar.Padding   = New-Object System.Windows.Forms.Padding(20, 0, 12, 0)
$root.Controls.Add($statusBar, 0, 1)

# Connection dot (painted circle)
$connDot = New-Object System.Windows.Forms.Panel
$connDot.Size     = New-Object System.Drawing.Size(12, 12)
$connDot.Location = New-Object System.Drawing.Point(2, 15)
$script:DotColor  = $Theme.DotRed
$connDot.Add_Paint({
    param($s, $e)
    $e.Graphics.SmoothingMode = 'AntiAlias'
    $b = New-Object System.Drawing.SolidBrush($script:DotColor)
    $e.Graphics.FillEllipse($b, 0, 0, 11, 11)
    $b.Dispose()
})
$statusBar.Controls.Add($connDot)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text      = 'Not connected'
$lblStatus.Font      = $FontUI
$lblStatus.AutoSize  = $true
$lblStatus.Location  = New-Object System.Drawing.Point(22, 13)
$statusBar.Controls.Add($lblStatus)

$btnConnect = New-Object System.Windows.Forms.Button
$btnConnect.Text   = 'Connect to Intune'
$btnConnect.Anchor = 'Top,Right'
Set-PrimaryButton -Btn $btnConnect -Back $Theme.Primary -Compact
$statusBar.Controls.Add($btnConnect)

# Right-align and vertically centre the Connect button within the status bar
function Update-ConnectButtonLayout {
    $y = [int](($statusBar.ClientSize.Height - $btnConnect.Height) / 2)
    if ($y -lt 0) { $y = 0 }
    $btnConnect.Location = New-Object System.Drawing.Point(($statusBar.ClientSize.Width - $btnConnect.Width - 16), $y)
}
$statusBar.Add_Resize({ Update-ConnectButtonLayout })
Update-ConnectButtonLayout

# ── TOOLBAR ──────────────────────────────────────────────────────────────────
$toolbar = New-Object System.Windows.Forms.Panel
$toolbar.Dock      = 'Fill'
$toolbar.BackColor = $Theme.ToolbarBg
$toolbar.Add_Paint({
    param($s, $e)
    $pen = New-Object System.Drawing.Pen($Theme.ToolbarLine, 1)
    $e.Graphics.DrawLine($pen, 0, ($s.ClientSize.Height - 1), $s.ClientSize.Width, ($s.ClientSize.Height - 1))
    $pen.Dispose()
})
$root.Controls.Add($toolbar, 0, 2)

$flow = New-Object System.Windows.Forms.FlowLayoutPanel
$flow.Dock          = 'Fill'
$flow.FlowDirection = 'LeftToRight'
$flow.WrapContents  = $false
$flow.Padding       = New-Object System.Windows.Forms.Padding(10, 0, 10, 0)
$toolbar.Controls.Add($flow)

$btnPull    = New-Object System.Windows.Forms.Button; $btnPull.Text    = 'Pull'
$btnTypes   = New-Object System.Windows.Forms.Button; $btnTypes.Text   = 'Content Types  ▾'
$btnExport  = New-Object System.Windows.Forms.Button; $btnExport.Text  = 'Export'
$btnImport  = New-Object System.Windows.Forms.Button; $btnImport.Text  = 'Import'
$btnFind    = New-Object System.Windows.Forms.Button; $btnFind.Text    = 'Find/Replace'
$btnBackup  = New-Object System.Windows.Forms.Button; $btnBackup.Text  = 'Backup'
$btnRestore = New-Object System.Windows.Forms.Button; $btnRestore.Text = 'Restore'
$btnApply   = New-Object System.Windows.Forms.Button; $btnApply.Text   = 'Apply'
$btnClear   = New-Object System.Windows.Forms.Button; $btnClear.Text   = 'Clear'

Set-PrimaryButton -Btn $btnPull  -Back $Theme.AccentBlue
Set-ToolButton    -Btn $btnTypes
Set-ToolButton    -Btn $btnExport
Set-ToolButton    -Btn $btnImport
Set-ToolButton    -Btn $btnFind
Set-ToolButton    -Btn $btnBackup
Set-ToolButton    -Btn $btnRestore
Set-PrimaryButton -Btn $btnApply -Back $Theme.Primary
Set-ToolButton    -Btn $btnClear

$chkDryRun = New-Object System.Windows.Forms.CheckBox
$chkDryRun.Text     = 'Dry run'
$chkDryRun.AutoSize = $true
$chkDryRun.Font     = $FontUI
$chkDryRun.Margin   = New-Object System.Windows.Forms.Padding(6, 17, 10, 0)

# Tooltips carry the detail the short labels drop
$tip = New-Object System.Windows.Forms.ToolTip
$tip.SetToolTip($btnPull,    'Pull the selected content types from Intune')
$tip.SetToolTip($btnTypes,   'Choose which Intune content types to pull')
$tip.SetToolTip($btnExport,  'Export the grid to a CSV file')
$tip.SetToolTip($btnImport,  'Import an edited CSV back into the grid')
$tip.SetToolTip($btnFind,    'Find & replace across New Name / New Description')
$tip.SetToolTip($btnBackup,  'Save a JSON snapshot of current names & descriptions')
$tip.SetToolTip($btnRestore, 'Restore names & descriptions from a JSON backup')
$tip.SetToolTip($btnApply,   'Write changed profiles to Intune')
$tip.SetToolTip($btnClear,   'Remove all rows from the grid')
$tip.SetToolTip($chkDryRun,  'Preview changes without writing anything to Intune')

# Initial enablement — set authoritatively by Set-ActionButtonsEnabled
$btnExport.Enabled  = $false
$btnFind.Enabled    = $false
$btnBackup.Enabled  = $false
$btnApply.Enabled   = $false
$btnRestore.Enabled = $false

# Grouped: [Pull · Content Types] | [Export / Import / Find] | [Backup / Restore] | [Apply / Dry run] | [Clear]
$flow.Controls.Add($btnPull)
$flow.Controls.Add($btnTypes)
Add-FlowSeparator -Flow $flow | Out-Null
$flow.Controls.Add($btnExport)
$flow.Controls.Add($btnImport)
$flow.Controls.Add($btnFind)
Add-FlowSeparator -Flow $flow | Out-Null
$flow.Controls.Add($btnBackup)
$flow.Controls.Add($btnRestore)
Add-FlowSeparator -Flow $flow | Out-Null
$flow.Controls.Add($btnApply)
$flow.Controls.Add($chkDryRun)
Add-FlowSeparator -Flow $flow | Out-Null
$flow.Controls.Add($btnClear)

# ── SPLIT: grid (top) + log (bottom) ─────────────────────────────────────────
$split = New-Object System.Windows.Forms.SplitContainer
$split.Dock        = 'Fill'
$split.Orientation = 'Horizontal'
$split.SplitterWidth = 6
$split.BackColor   = $Theme.Border
$root.Controls.Add($split, 0, 3)

# DataGridView
$grid = New-Object System.Windows.Forms.DataGridView
$grid.Dock                       = 'Fill'
$grid.AllowUserToAddRows         = $false
$grid.AllowUserToDeleteRows      = $false
$grid.AllowUserToResizeRows      = $false
$grid.RowHeadersVisible          = $false
$grid.SelectionMode              = 'FullRowSelect'
$grid.EditMode                   = 'EditOnKeystrokeOrF2'
$grid.AutoSizeColumnsMode        = 'Fill'
$grid.BorderStyle                = 'None'
$grid.BackgroundColor            = $Theme.White
$grid.GridColor                  = $Theme.GridLine
$grid.EnableHeadersVisualStyles  = $false
$grid.ColumnHeadersHeightSizeMode = 'DisableResizing'
$grid.ColumnHeadersHeight        = 32
$grid.ColumnHeadersBorderStyle   = 'Single'
$grid.ColumnHeadersDefaultCellStyle.BackColor  = $Theme.GridHeaderBg
$grid.ColumnHeadersDefaultCellStyle.ForeColor  = $Theme.GridHeaderFg
$grid.ColumnHeadersDefaultCellStyle.Font       = $FontUIBold
$grid.ColumnHeadersDefaultCellStyle.Padding    = New-Object System.Windows.Forms.Padding(6, 0, 0, 0)
$grid.AlternatingRowsDefaultCellStyle.BackColor = $Theme.GridAltRow
$grid.DefaultCellStyle.SelectionBackColor      = $Theme.GridSelBg
$grid.DefaultCellStyle.SelectionForeColor      = $Theme.GridSelFg
$grid.DefaultCellStyle.Padding                 = New-Object System.Windows.Forms.Padding(4, 0, 0, 0)
$grid.RowTemplate.Height = 26

function Add-GridColumn {
    param([string]$Name, [string]$Header, [bool]$ReadOnly, [int]$Weight, [bool]$Visible = $true)
    $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $col.Name        = $Name
    $col.HeaderText  = $Header
    $col.ReadOnly    = $ReadOnly
    $col.FillWeight  = $Weight
    $col.Visible     = $Visible
    $col.SortMode    = 'NotSortable'
    if ($ReadOnly) { $col.DefaultCellStyle.ForeColor = C '#555555' }
    $grid.Columns.Add($col) | Out-Null
}

Add-GridColumn -Name 'ProfileId'          -Header 'ProfileId'        -ReadOnly $true  -Weight 1  -Visible $false
Add-GridColumn -Name 'ProfileType'        -Header 'Type'             -ReadOnly $true  -Weight 14
Add-GridColumn -Name 'CurrentName'        -Header 'Current Name'     -ReadOnly $true  -Weight 22
Add-GridColumn -Name 'NewName'            -Header 'New Name'         -ReadOnly $false -Weight 22
Add-GridColumn -Name 'CurrentDescription' -Header 'Current Description' -ReadOnly $true -Weight 21
Add-GridColumn -Name 'NewDescription'    -Header 'New Description'  -ReadOnly $false -Weight 21

$split.Panel1.Controls.Add($grid)

# Log area (Panel2): header strip + Consolas box
$logHeader = New-Object System.Windows.Forms.Panel
$logHeader.Dock      = 'Top'
$logHeader.Height    = 26
$logHeader.BackColor = $Theme.White

$lblLog = New-Object System.Windows.Forms.Label
$lblLog.Text      = 'Activity Log'
$lblLog.Font      = $FontUIBold
$lblLog.AutoSize  = $true
$lblLog.Location  = New-Object System.Drawing.Point(4, 5)
$logHeader.Controls.Add($lblLog)

$btnClearLog = New-Object System.Windows.Forms.Button
$btnClearLog.Text   = 'Clear'
$btnClearLog.Height = 22
$btnClearLog.Anchor = 'Top,Right'
Set-ToolButton -Btn $btnClearLog
$btnClearLog.AutoSize = $false
$btnClearLog.Width    = 60
$logHeader.Controls.Add($btnClearLog)
$logHeader.Add_Resize({ $btnClearLog.Location = New-Object System.Drawing.Point(($logHeader.ClientSize.Width - $btnClearLog.Width - 4), 2) }.GetNewClosure())

$logBox = New-Object System.Windows.Forms.RichTextBox
$logBox.Dock        = 'Fill'
$logBox.ReadOnly    = $true
$logBox.BackColor   = $Theme.White
$logBox.BorderStyle = 'None'
$logBox.Font        = $FontMono
$logBox.WordWrap    = $true

$split.Panel2.Controls.Add($logBox)
$split.Panel2.Controls.Add($logHeader)
$split.Panel2.Padding = New-Object System.Windows.Forms.Padding(2, 0, 2, 2)

# ── FOOTER ───────────────────────────────────────────────────────────────────
$footer = New-Object System.Windows.Forms.Panel
$footer.Dock      = 'Fill'
$footer.BackColor = $Theme.FooterBg
$footer.Add_Paint({
    param($s, $e)
    $pen = New-Object System.Drawing.Pen($Theme.Border, 1)
    $e.Graphics.DrawLine($pen, 0, 0, $s.ClientSize.Width, 0)
    $pen.Dispose()
})
$root.Controls.Add($footer, 0, 4)

$lblFooter = New-Object System.Windows.Forms.Label
$lblFooter.Text      = 'Ready'
$lblFooter.Font      = $FontSub
$lblFooter.ForeColor = $Theme.FooterText
$lblFooter.AutoSize  = $true
$lblFooter.Location  = New-Object System.Drawing.Point(16, 7)
$footer.Controls.Add($lblFooter)

$lblFooterR = New-Object System.Windows.Forms.Label
$lblFooterR.Text      = 'Intune Profile Manager — modernworkspacehub.com — Provided without warranty'
$lblFooterR.Font      = New-Object System.Drawing.Font('Segoe UI', 8)
$lblFooterR.ForeColor = $Theme.FooterFaint
$lblFooterR.AutoSize  = $true
$lblFooterR.Anchor    = 'Top,Right'
$footer.Controls.Add($lblFooterR)
$footer.Add_Resize({ $lblFooterR.Location = New-Object System.Drawing.Point(($footer.ClientSize.Width - $lblFooterR.Width - 16), 7) }.GetNewClosure())
$lblFooterR.Location = New-Object System.Drawing.Point(($footer.ClientSize.Width - $lblFooterR.Width - 16), 7)

# Splitter distance must be set after the form has a size
$form.Add_Shown({
    try { $split.SplitterDistance = [int]($split.Height * 0.62) } catch {}
    $logBox.Focus() | Out-Null
})
#endregion

# ─────────────────────────────────────────────────────────────────────────────
#region Helper functions
# ─────────────────────────────────────────────────────────────────────────────
function Write-Log {
    param([string]$Text, [ValidateSet('Info','OK','Warn','Fail')][string]$Level = 'Info')
    $prefix, $color = switch ($Level) {
        'OK'   { '[OK]   ', $Theme.LogOk }
        'Warn' { '[WARN] ', $Theme.LogWarn }
        'Fail' { '[FAIL] ', $Theme.LogFail }
        default{ '[INFO] ', $Theme.LogInfo }
    }
    $line = "$(Get-Date -Format 'HH:mm:ss')  $prefix $Text`n"
    $logBox.SelectionStart  = $logBox.TextLength
    $logBox.SelectionLength = 0
    $logBox.SelectionColor  = $color
    $logBox.AppendText($line)
    $logBox.SelectionColor  = $logBox.ForeColor
    $logBox.ScrollToCaret()

    # Persist to the rolling daily log file in the repo (best-effort — never break the UI)
    try {
        $fileLine = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  $prefix $Text"
        Add-Content -Path $script:LogFile -Value $fileLine -Encoding UTF8
    } catch {}

    [System.Windows.Forms.Application]::DoEvents()
}

function Set-Footer { param([string]$Text) $lblFooter.Text = $Text; [System.Windows.Forms.Application]::DoEvents() }

function Set-ConnectedState {
    param([bool]$On, [string]$Account = '', [string]$Tenant = '')
    $script:Connected = $On
    if ($On) {
        $script:DotColor   = $Theme.DotGreen
        $lblStatus.Text    = "Connected as $Account   •   Tenant: $Tenant"
        $btnConnect.Text   = 'Disconnect'
        Set-PrimaryButton -Btn $btnConnect -Back $Theme.PrimaryDark -Compact
    } else {
        $script:DotColor   = $Theme.DotRed
        $lblStatus.Text    = 'Not connected'
        $btnConnect.Text   = 'Connect to Intune'
        Set-PrimaryButton -Btn $btnConnect -Back $Theme.Primary -Compact
    }
    Update-ConnectButtonLayout
    $connDot.Invalidate()
}

# Robust string compare treating $null and '' as equal
function Test-Differs { param($A, $B) return (([string]$A) -ne ([string]$B)) }

function Test-RowChanged {
    param($Row)
    if ($Row.IsNewRow) { return $false }
    return (Test-Differs $Row.Cells['CurrentName'].Value        $Row.Cells['NewName'].Value) -or
           (Test-Differs $Row.Cells['CurrentDescription'].Value $Row.Cells['NewDescription'].Value)
}

# Re-colour the editable cells of one row to flag pending changes
function Update-RowHighlight {
    param($Row)
    if ($Row.IsNewRow) { return }
    $nameChanged = Test-Differs $Row.Cells['CurrentName'].Value        $Row.Cells['NewName'].Value
    $descChanged = Test-Differs $Row.Cells['CurrentDescription'].Value $Row.Cells['NewDescription'].Value

    $nameCell = $Row.Cells['NewName']
    if ($nameChanged) { $nameCell.Style.BackColor = $Theme.ChangedBg; $nameCell.Style.ForeColor = $Theme.ChangedFg }
    else              { $nameCell.Style.BackColor = [System.Drawing.Color]::Empty; $nameCell.Style.ForeColor = [System.Drawing.Color]::Empty }

    $descCell = $Row.Cells['NewDescription']
    if ($descChanged) { $descCell.Style.BackColor = $Theme.ChangedBg; $descCell.Style.ForeColor = $Theme.ChangedFg }
    else              { $descCell.Style.BackColor = [System.Drawing.Color]::Empty; $descCell.Style.ForeColor = [System.Drawing.Color]::Empty }
}

function Update-AllHighlights {
    foreach ($r in $grid.Rows) { Update-RowHighlight -Row $r }
}

function Get-ChangedRowCount {
    $n = 0
    foreach ($r in $grid.Rows) { if (Test-RowChanged -Row $r) { $n++ } }
    return $n
}

function Set-ActionButtonsEnabled {
    $has = $grid.Rows.Count -gt 0
    $btnExport.Enabled  = $has
    $btnFind.Enabled    = $has
    $btnBackup.Enabled  = $has
    $btnApply.Enabled   = $has -and $script:Connected
    $btnRestore.Enabled = $script:Connected
}

# Graph collection getter with automatic @odata.nextLink paging
function Get-GraphCollection {
    param([Parameter(Mandatory)][string]$Uri)
    $items = [System.Collections.Generic.List[object]]::new()
    $next  = $Uri
    do {
        $resp = Invoke-MgGraphRequest -Method GET -Uri $next -OutputType PSObject
        if ($resp.value) { $items.AddRange([object[]]$resp.value) }
        $next = $resp.'@odata.nextLink'
    } while ($next)
    return $items
}

# Map a deviceConfiguration @odata.type to a friendly template kind shown in brackets,
# e.g. "#microsoft.graph.windows81VpnConfiguration" -> "VPN".
function Get-TemplateSubtype {
    param([string]$ODataType)
    if (-not $ODataType) { return $null }
    $t = ($ODataType.TrimStart('#')) -replace '^microsoft\.graph\.', ''
    switch -Regex ($t.ToLower()) {
        'domainjoin'                 { return 'Domain join' }
        'scepcertificate'            { return 'SCEP certificate' }
        'pkcscertificate'            { return 'PKCS certificate' }
        'trustedrootcertificate'     { return 'Trusted certificate' }
        'vpnconfiguration'           { return 'VPN' }
        'wifi'                       { return 'Wi-Fi' }
        'wirednetwork'               { return 'Wired network' }
        'healthmonitoring'           { return 'Health monitoring' }
        'endpointprotection'         { return 'Endpoint protection' }
        'identityprotection'         { return 'Identity protection' }
        'advancedthreatprotection'   { return 'Defender for Endpoint' }
        'devicefirmware'             { return 'DFCI' }
        'kiosk'                      { return 'Kiosk' }
        'editionupgrade'             { return 'Edition upgrade' }
        'easemail|emailprofile'      { return 'Email' }
        'customconfiguration'        { return 'Custom (OMA-URI)' }
        'deliveryoptimization'       { return 'Delivery optimization' }
        'networkboundary'            { return 'Network boundary' }
        'secureassessment'           { return 'Secure assessment' }
        'sharedpc'                   { return 'Shared multi-user' }
        'updateforbusiness'          { return 'Windows Update ring' }
        'teamgeneral'                { return 'Surface Hub' }
        'generalconfiguration|generaldeviceconfiguration' { return 'Device restrictions' }
        default {
            # Fall back to a tidied concrete type name (strip suffix, split camelCase)
            $name = $t -replace '(Configuration|Profile)$', ''
            $name = $name -creplace '([a-z0-9])([A-Z])', '$1 $2'
            if ($name) { return $name.Trim() } else { return $null }
        }
    }
}

# The grid/CSV ProfileType may carry a subtype suffix "Base (Subtype)"; routing needs the base.
function Get-BaseProfileType {
    param([string]$Value)
    if (-not $Value) { return $Value }
    $i = $Value.IndexOf(' (')
    if ($i -ge 0) { return $Value.Substring(0, $i) }
    return $Value
}

function Add-ProfileRow {
    param([string]$Id, [string]$Type, [string]$CurName, [string]$CurDesc, [string]$NewName, [string]$NewDesc)
    $idx = $grid.Rows.Add()
    $row = $grid.Rows[$idx]
    $row.Cells['ProfileId'].Value          = $Id
    $row.Cells['ProfileType'].Value        = $Type
    $row.Cells['CurrentName'].Value        = $CurName
    $row.Cells['NewName'].Value            = $NewName
    $row.Cells['CurrentDescription'].Value = $CurDesc
    $row.Cells['NewDescription'].Value     = $NewDesc
}
#endregion

# ─────────────────────────────────────────────────────────────────────────────
#region Connect / Disconnect
# ─────────────────────────────────────────────────────────────────────────────
$btnConnect.Add_Click({
    if ($script:Connected) {
        try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}
        Set-ConnectedState -On $false
        Set-ActionButtonsEnabled
        Set-Footer 'Disconnected'
        Write-Log 'Signed out of Microsoft Graph.' 'Info'
        return
    }

    Set-Footer 'Connecting to Microsoft Graph...'
    try {
        Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    }
    catch {
        Write-Log 'Microsoft Graph SDK not found.' 'Fail'
        [System.Windows.Forms.MessageBox]::Show(
            "The Microsoft Graph PowerShell SDK is required but not installed.`n`n" +
            "Install it with:`n`n    Install-Module Microsoft.Graph.Authentication -Scope CurrentUser`n`n" +
            "Then restart this tool.",
            'Graph SDK Missing', 'OK', 'Warning') | Out-Null
        Set-Footer 'Ready'
        return
    }

    try {
        Write-Log "Opening sign-in — requesting $($RequiredScopes.Count) delegated scopes..." 'Info'
        Connect-MgGraph -Scopes $RequiredScopes -NoWelcome -ErrorAction Stop | Out-Null
        $ctx = Get-MgContext
        if (-not $ctx) { throw 'No Graph context returned after sign-in.' }

        # Note any requested scopes that weren't granted — those content types will be skipped on pull
        $missing = @($RequiredScopes | Where-Object { $ctx.Scopes -notcontains $_ })
        if ($missing.Count -gt 0) {
            Write-Log "Not consented: $($missing -join ', '). Content types needing these will be skipped." 'Warn'
        }

        Set-ConnectedState -On $true -Account $ctx.Account -Tenant $ctx.TenantId
        Set-ActionButtonsEnabled
        Write-Log "Connected to tenant $($ctx.TenantId) as $($ctx.Account)." 'OK'
        Set-Footer 'Connected. Choose Content Types if needed, then click Pull.'
    }
    catch {
        Set-ConnectedState -On $false
        Write-Log "Connection failed: $($_.Exception.Message)" 'Fail'
        [System.Windows.Forms.MessageBox]::Show(
            "Could not connect to Microsoft Graph:`n`n$($_.Exception.Message)`n`n" +
            "Check your network, account permissions, and that you consented to the requested scope.",
            'Connection Failed', 'OK', 'Error') | Out-Null
        Set-Footer 'Ready'
    }
})
#endregion

# ─────────────────────────────────────────────────────────────────────────────
#region Pull profiles
# ─────────────────────────────────────────────────────────────────────────────
$btnPull.Add_Click({
    if (-not $script:Connected) {
        [System.Windows.Forms.MessageBox]::Show('Please connect to Intune first.', 'Not Connected', 'OK', 'Warning') | Out-Null
        return
    }
    if ($grid.Rows.Count -gt 0) {
        $ans = [System.Windows.Forms.MessageBox]::Show(
            'Pulling will replace the current grid contents. Any unsaved edits will be lost. Continue?',
            'Replace Grid?', 'YesNo', 'Question')
        if ($ans -ne 'Yes') { return }
    }

    # Only the content types ticked in the Content Types picker
    $typesToPull = @($ContentTypes.Keys | Where-Object { $global:IPMSelectedContentTypes -contains $_ })
    if ($typesToPull.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('No content types selected. Click "Content Types" and tick at least one.', 'Nothing Selected', 'OK', 'Warning') | Out-Null
        return
    }

    $btnPull.Enabled = $false
    Set-Footer 'Pulling selected content types...'
    $script:Loading = $true
    $grid.Rows.Clear()
    $script:OdataMap.Clear()
    $total = 0; $anyFailed = $false
    Write-Log "Pulling $($typesToPull.Count) content type$(if ($typesToPull.Count -ne 1){'s'}): $($typesToPull -join ', ')" 'Info'

    foreach ($type in $typesToPull) {
        $kind = $ContentTypes[$type]
        $count = 0
        try {
            Write-Log "Fetching $type..." 'Info'
            $uri  = "$($kind.Base)/$($kind.Collection)?`$select=$($kind.Select)&`$top=100"
            $items = Get-GraphCollection -Uri $uri
            foreach ($p in $items) {
                $nm = [string]$p.($kind.NameProp)
                $ds = [string]$p.description
                $label = $type
                # Polymorphic types carry a derived @odata.type — capture it for the PATCH, and
                # (for Device Configuration) show the specific template kind, e.g. "(VPN)".
                if ($kind.NeedsODataType -and $p.'@odata.type') {
                    $script:OdataMap[$p.id] = $p.'@odata.type'
                    if ($kind.Subtype) {
                        $sub = Get-TemplateSubtype $p.'@odata.type'
                        if ($sub) { $label = "$type ($sub)" }
                    }
                }
                Add-ProfileRow -Id $p.id -Type $label -CurName $nm -CurDesc $ds -NewName $nm -NewDesc $ds
                $count++
            }
            Write-Log "Loaded $count $type item$(if ($count -ne 1){'s'})." 'OK'
            $total += $count
        }
        catch {
            $anyFailed = $true
            $msg = $_.Exception.Message
            if ($msg -match '403|Forbidden') {
                Write-Log "Skipped $type — access denied (role/scope not consented for this type)." 'Warn'
            } else {
                Write-Log "Failed to load $type`: $msg" 'Fail'
            }
        }
    }

    $script:Loading = $false
    Update-AllHighlights
    Set-ActionButtonsEnabled
    $btnPull.Enabled = $true

    Write-Log "Pull complete — $total item$(if ($total -ne 1){'s'}) total across $($typesToPull.Count) content type$(if ($typesToPull.Count -ne 1){'s'})." $(if ($anyFailed) { 'Warn' } else { 'OK' })
    Set-Footer "$total items loaded. Edit New Name / New Description, or Export to CSV."
})
#endregion

# ─────────────────────────────────────────────────────────────────────────────
#region Export CSV
# ─────────────────────────────────────────────────────────────────────────────
$btnExport.Add_Click({
    if ($grid.Rows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('Nothing to export — pull profiles first.', 'Empty Grid', 'OK', 'Information') | Out-Null
        return
    }

    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter   = 'CSV files (*.csv)|*.csv'
    $dlg.FileName = "IntuneProfiles_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $dlg.Title    = 'Export profiles to CSV'
    if ($dlg.ShowDialog() -ne 'OK') { return }

    try {
        $rows = foreach ($r in $grid.Rows) {
            if ($r.IsNewRow) { continue }
            [PSCustomObject]@{
                ProfileId          = [string]$r.Cells['ProfileId'].Value
                ProfileType        = [string]$r.Cells['ProfileType'].Value
                CurrentName        = [string]$r.Cells['CurrentName'].Value
                NewName            = [string]$r.Cells['NewName'].Value
                CurrentDescription = [string]$r.Cells['CurrentDescription'].Value
                NewDescription     = [string]$r.Cells['NewDescription'].Value
            }
        }
        $rows | Export-Csv -Path $dlg.FileName -NoTypeInformation -Encoding UTF8
        Write-Log "Exported $(@($rows).Count) profiles to: $($dlg.FileName)" 'OK'
        Set-Footer "Exported to $($dlg.FileName)"
        # Open the CSV in the default handler (usually Excel)
        Start-Process -FilePath $dlg.FileName
    }
    catch {
        Write-Log "Export failed: $($_.Exception.Message)" 'Fail'
        [System.Windows.Forms.MessageBox]::Show("Could not export CSV:`n`n$($_.Exception.Message)", 'Export Failed', 'OK', 'Error') | Out-Null
    }
})
#endregion

# ─────────────────────────────────────────────────────────────────────────────
#region Import CSV
# ─────────────────────────────────────────────────────────────────────────────
$btnImport.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = 'CSV files (*.csv)|*.csv'
    $dlg.Title  = 'Import edited profile CSV'
    if ($dlg.ShowDialog() -ne 'OK') { return }

    try {
        $data = Import-Csv -Path $dlg.FileName
    }
    catch {
        Write-Log "Import failed: $($_.Exception.Message)" 'Fail'
        [System.Windows.Forms.MessageBox]::Show("Could not read CSV:`n`n$($_.Exception.Message)", 'Import Failed', 'OK', 'Error') | Out-Null
        return
    }

    if (-not $data -or @($data).Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('The CSV file is empty.', 'Empty CSV', 'OK', 'Warning') | Out-Null
        return
    }

    # Validate required columns
    $required = @('ProfileId','ProfileType','CurrentName','NewName','CurrentDescription','NewDescription')
    $present  = @($data[0].PSObject.Properties.Name)
    $missing  = $required | Where-Object { $_ -notin $present }
    if ($missing.Count -gt 0) {
        Write-Log "Import rejected — missing columns: $($missing -join ', ')" 'Fail'
        [System.Windows.Forms.MessageBox]::Show(
            "The CSV is missing required column(s):`n`n  $($missing -join "`n  ")`n`n" +
            "Expected columns:`n  $($required -join ', ')",
            'Invalid CSV', 'OK', 'Error') | Out-Null
        return
    }

    $script:Loading = $true
    $grid.Rows.Clear()
    $script:OdataMap.Clear()
    $loaded = 0; $skipped = 0; $lineNo = 1
    foreach ($row in $data) {
        $lineNo++
        $id   = [string]$row.ProfileId
        $type = [string]$row.ProfileType
        if ([string]::IsNullOrWhiteSpace($id) -or [string]::IsNullOrWhiteSpace($type)) {
            Write-Log "Skipped CSV line $lineNo — blank ProfileId or ProfileType." 'Warn'
            $skipped++; continue
        }
        if (-not $ContentTypes.Contains((Get-BaseProfileType $type))) {
            Write-Log "Skipped CSV line $lineNo — unknown ProfileType '$type'." 'Warn'
            $skipped++; continue
        }
        Add-ProfileRow -Id $id -Type $type `
            -CurName $row.CurrentName -CurDesc $row.CurrentDescription `
            -NewName $row.NewName     -NewDesc $row.NewDescription
        $loaded++
    }
    $script:Loading = $false
    Update-AllHighlights
    Set-ActionButtonsEnabled

    $changed = Get-ChangedRowCount
    Write-Log "Imported $loaded row$(if ($loaded -ne 1){'s'}) ($changed with changes, $skipped skipped) from $(Split-Path $dlg.FileName -Leaf)." 'OK'
    Set-Footer "Imported $loaded profiles — $changed pending change$(if ($changed -ne 1){'s'})."
    if ($skipped -gt 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "$loaded row(s) imported.`n$skipped row(s) skipped (see Activity Log for details).",
            'Import Complete', 'OK', 'Warning') | Out-Null
    }
})
#endregion

# ─────────────────────────────────────────────────────────────────────────────
#region Apply changes
# ─────────────────────────────────────────────────────────────────────────────
# Resolve a polymorphic item's derived @odata.type (required on PATCH for some types).
# Prefers the value captured during pull; otherwise GETs the single item from its endpoint.
function Get-ContentODataType {
    param([string]$Base, [string]$Collection, [string]$Id)
    if ($script:OdataMap.ContainsKey($Id) -and $script:OdataMap[$Id]) { return $script:OdataMap[$Id] }
    $item = Invoke-MgGraphRequest -Method GET -Uri "$Base/$Collection/$Id" -OutputType PSObject
    $odt  = $item.'@odata.type'
    if ($odt) { $script:OdataMap[$Id] = $odt }
    return $odt
}

# Single PATCH path used by both Apply and Restore — routes by content type via $ContentTypes.
# Only ever touches the display name and description; nothing else about the item is sent.
function Invoke-ProfilePatch {
    param([string]$Type, [string]$Id, [string]$Name, [string]$Desc)
    $kind = $ContentTypes[(Get-BaseProfileType $Type)]
    if (-not $kind) { throw "Unknown content type '$Type'." }
    $body = @{ description = $Desc }
    $body[$kind.PatchNameProp] = $Name
    if ($kind.NeedsODataType) {
        $odt = Get-ContentODataType -Base $kind.Base -Collection $kind.Collection -Id $Id
        if (-not $odt) { throw 'Could not determine @odata.type for this item.' }
        $body['@odata.type'] = $odt
    }
    Invoke-MgGraphRequest -Method PATCH -Uri "$($kind.Base)/$($kind.Collection)/$Id" -Body $body | Out-Null
}

# Snapshot the CURRENT (live) name + description of every grid row to a timestamped JSON
# in .\Backups so changes can be reverted via Restore JSON. Returns the file path (or $null if empty).
function New-ProfileBackup {
    param([switch]$Auto)
    $snapshot = foreach ($r in $grid.Rows) {
        if ($r.IsNewRow) { continue }
        [PSCustomObject]@{
            ProfileId   = [string]$r.Cells['ProfileId'].Value
            ProfileType = [string]$r.Cells['ProfileType'].Value
            Name        = [string]$r.Cells['CurrentName'].Value
            Description = [string]$r.Cells['CurrentDescription'].Value
        }
    }
    $snapshot = @($snapshot)
    if ($snapshot.Count -eq 0) { return $null }

    $tenant = ''
    try { $tenant = (Get-MgContext).TenantId } catch {}
    $tag  = if ($Auto) { 'auto-preapply' } else { 'manual' }
    $path = Join-Path $script:BackupDir "ProfileBackup_$(Get-Date -Format 'yyyyMMdd_HHmmss')_$tag.json"
    [PSCustomObject]@{
        CreatedUtc   = (Get-Date).ToUniversalTime().ToString('s') + 'Z'
        Tenant       = $tenant
        ProfileCount = $snapshot.Count
        Profiles     = $snapshot
    } | ConvertTo-Json -Depth 6 | Set-Content -Path $path -Encoding UTF8
    return $path
}

$btnApply.Add_Click({
    if (-not $script:Connected) {
        [System.Windows.Forms.MessageBox]::Show('Please connect to Intune first.', 'Not Connected', 'OK', 'Warning') | Out-Null
        return
    }

    # Commit any in-progress cell edit before reading values
    try { $grid.EndEdit() | Out-Null } catch {}

    $changedRows = @($grid.Rows | Where-Object { -not $_.IsNewRow -and (Test-RowChanged -Row $_) })
    if ($changedRows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('No changes detected — every New value matches its Current value.', 'Nothing to Apply', 'OK', 'Information') | Out-Null
        return
    }

    $dry = $chkDryRun.Checked
    $verb = if ($dry) { 'Preview' } else { 'Apply' }
    $ans = [System.Windows.Forms.MessageBox]::Show(
        "$($changedRows.Count) profile(s) have changes.`n`n" +
        $(if ($dry) { "DRY RUN: nothing will actually be modified in Intune." }
          else      { "These changes will be written to Intune via Microsoft Graph.`nOnly display name and description are modified." }) +
        "`n`nProceed?",
        "$verb Changes", 'YesNo', $(if ($dry) { 'Information' } else { 'Warning' }))
    if ($ans -ne 'Yes') { return }

    $btnApply.Enabled = $false; $btnPull.Enabled = $false
    $ok = 0; $fail = 0; $i = 0

    # Automatic safety backup of current values before any real write, so changes can be reverted
    if (-not $dry) {
        try {
            $autoBackup = New-ProfileBackup -Auto
            if ($autoBackup) { Write-Log "Safety backup written: $autoBackup" 'Info' }
        } catch {
            Write-Log "Could not write safety backup: $($_.Exception.Message)" 'Warn'
        }
    }

    Write-Log "─── ${verb}: $($changedRows.Count) profile(s)$(if ($dry){' [DRY RUN]'}) ───" 'Info'

    foreach ($row in $changedRows) {
        $i++
        $id      = [string]$row.Cells['ProfileId'].Value
        $type    = [string]$row.Cells['ProfileType'].Value
        $newName = [string]$row.Cells['NewName'].Value
        $newDesc = [string]$row.Cells['NewDescription'].Value
        $curName = [string]$row.Cells['CurrentName'].Value
        Set-Footer "$verb $i of $($changedRows.Count): $newName"

        # Guard: don't let a profile be blanked out
        if ([string]::IsNullOrWhiteSpace($newName)) {
            Write-Log "[$i/$($changedRows.Count)] Skipped '$curName' — New Name is blank." 'Warn'
            $fail++; continue
        }

        try {
            if ($dry) {
                $bits = @()
                if (Test-Differs $row.Cells['CurrentName'].Value $newName) { $bits += "name → '$newName'" }
                if (Test-Differs $row.Cells['CurrentDescription'].Value $newDesc) { $bits += 'description changed' }
                Write-Log "[$i/$($changedRows.Count)] [DRY RUN] $type '$curName' : $($bits -join ', ')" 'Info'
                $ok++
                continue
            }

            Invoke-ProfilePatch -Type $type -Id $id -Name $newName -Desc $newDesc

            # Success — the New values are now the Current values
            $row.Cells['CurrentName'].Value        = $newName
            $row.Cells['CurrentDescription'].Value = $newDesc
            Update-RowHighlight -Row $row
            Write-Log "[$i/$($changedRows.Count)] Updated $type '$newName'." 'OK'
            $ok++
        }
        catch {
            $msg = $_.Exception.Message
            $hint = ''
            if ($msg -match '403|Forbidden') { $hint = ' (permission denied — check DeviceManagementConfiguration.ReadWrite.All consent)' }
            elseif ($msg -match '404|NotFound') { $hint = ' (profile no longer exists — try Pull Profiles again)' }
            Write-Log "[$i/$($changedRows.Count)] FAILED '$curName': $msg$hint" 'Fail'
            $fail++
        }
    }

    $summary = "$verb complete — $ok succeeded, $fail failed."
    Write-Log $summary $(if ($fail -gt 0) { 'Warn' } else { 'OK' })
    Set-Footer $summary
    $btnApply.Enabled = $true; $btnPull.Enabled = $true
    Set-ActionButtonsEnabled

    [System.Windows.Forms.MessageBox]::Show(
        $summary + $(if ($dry) { "`n`n(Dry run — nothing was actually changed.)" } else { '' }),
        "$verb Results", 'OK', $(if ($fail -gt 0) { 'Warning' } else { 'Information' })) | Out-Null
})
#endregion

# ─────────────────────────────────────────────────────────────────────────────
#region Backup / Restore (JSON)
# ─────────────────────────────────────────────────────────────────────────────
$btnBackup.Add_Click({
    if ($grid.Rows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('Nothing to back up — pull profiles first.', 'Empty Grid', 'OK', 'Information') | Out-Null
        return
    }
    try {
        $path = New-ProfileBackup
        if ($path) {
            Write-Log "Backup written: $path" 'OK'
            Set-Footer "Backup saved to $path"
            [System.Windows.Forms.MessageBox]::Show(
                "Saved a JSON snapshot of all current profile names and descriptions to:`n`n$path`n`n" +
                'Use "Restore JSON" to revert to these values later.',
                'Backup Complete', 'OK', 'Information') | Out-Null
        }
    }
    catch {
        Write-Log "Backup failed: $($_.Exception.Message)" 'Fail'
        [System.Windows.Forms.MessageBox]::Show("Could not write backup:`n`n$($_.Exception.Message)", 'Backup Failed', 'OK', 'Error') | Out-Null
    }
})

$btnRestore.Add_Click({
    if (-not $script:Connected) {
        [System.Windows.Forms.MessageBox]::Show('Please connect to Intune first.', 'Not Connected', 'OK', 'Warning') | Out-Null
        return
    }

    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter           = 'JSON backups (*.json)|*.json'
    $dlg.Title            = 'Restore profiles from a JSON backup'
    if (Test-Path $script:BackupDir) { $dlg.InitialDirectory = $script:BackupDir }
    if ($dlg.ShowDialog() -ne 'OK') { return }

    try {
        $data = Get-Content -Path $dlg.FileName -Raw | ConvertFrom-Json
    }
    catch {
        Write-Log "Restore failed — could not read backup: $($_.Exception.Message)" 'Fail'
        [System.Windows.Forms.MessageBox]::Show("Could not read backup JSON:`n`n$($_.Exception.Message)", 'Restore Failed', 'OK', 'Error') | Out-Null
        return
    }

    # Accept either the wrapped format ({ Profiles: [...] }) or a bare array
    $profiles = if ($data.PSObject.Properties.Name -contains 'Profiles') { @($data.Profiles) } else { @($data) }
    $profiles = @($profiles | Where-Object { $_.ProfileId -and $_.ProfileType })
    if ($profiles.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('No valid profiles found in that backup file.', 'Nothing to Restore', 'OK', 'Warning') | Out-Null
        return
    }

    $dry  = $chkDryRun.Checked
    $verb = if ($dry) { 'Preview restore of' } else { 'Restore' }
    $ans = [System.Windows.Forms.MessageBox]::Show(
        "$verb $($profiles.Count) profile(s) from:`n$(Split-Path $dlg.FileName -Leaf)`n`n" +
        $(if ($dry) { 'DRY RUN: nothing will actually be changed.' }
          else      { 'Each profile''s name and description will be overwritten in Intune with the backed-up values.' }) +
        "`n`nProceed?",
        'Confirm Restore', 'YesNo', $(if ($dry) { 'Information' } else { 'Warning' }))
    if ($ans -ne 'Yes') { return }

    $btnRestore.Enabled = $false; $btnApply.Enabled = $false; $btnPull.Enabled = $false
    $ok = 0; $fail = 0; $i = 0
    Write-Log "─── ${verb} $($profiles.Count) profile(s) from backup$(if ($dry){' [DRY RUN]'}) ───" 'Info'

    foreach ($p in $profiles) {
        $i++
        $name = [string]$p.Name
        $desc = [string]$p.Description
        Set-Footer "$verb $i of $($profiles.Count): $name"
        if ([string]::IsNullOrWhiteSpace($name)) {
            Write-Log "[$i/$($profiles.Count)] Skipped $($p.ProfileType) $($p.ProfileId) — backed-up name is blank." 'Warn'
            $fail++; continue
        }
        try {
            if ($dry) {
                Write-Log "[$i/$($profiles.Count)] [DRY RUN] Would restore $($p.ProfileType) '$name'." 'Info'
            } else {
                Invoke-ProfilePatch -Type ([string]$p.ProfileType) -Id ([string]$p.ProfileId) -Name $name -Desc $desc
                Write-Log "[$i/$($profiles.Count)] Restored $($p.ProfileType) '$name'." 'OK'
            }
            $ok++
        }
        catch {
            Write-Log "[$i/$($profiles.Count)] FAILED to restore '$name': $($_.Exception.Message)" 'Fail'
            $fail++
        }
    }

    $summary = "Restore complete — $ok succeeded, $fail failed."
    Write-Log $summary $(if ($fail -gt 0) { 'Warn' } else { 'OK' })
    Set-Footer $summary
    Set-ActionButtonsEnabled
    $btnPull.Enabled = $true
    [System.Windows.Forms.MessageBox]::Show(
        $summary + $(if ($dry) { "`n`n(Dry run — nothing was actually changed.)" } else { "`n`nTip: click Pull Profiles to refresh the grid with the restored values." }),
        'Restore Results', 'OK', $(if ($fail -gt 0) { 'Warning' } else { 'Information' })) | Out-Null
})
#endregion

# ─────────────────────────────────────────────────────────────────────────────
#region Grid edit / misc events
# ─────────────────────────────────────────────────────────────────────────────
$grid.Add_CellEndEdit({
    param($s, $e)
    if ($script:Loading) { return }
    Update-RowHighlight -Row $grid.Rows[$e.RowIndex]
})

$btnClear.Add_Click({
    if ($grid.Rows.Count -eq 0) { return }
    $ans = [System.Windows.Forms.MessageBox]::Show('Clear all rows from the grid?', 'Clear Grid', 'YesNo', 'Question')
    if ($ans -eq 'Yes') {
        $grid.Rows.Clear()
        $script:OdataMap.Clear()
        Set-ActionButtonsEnabled
        Write-Log 'Grid cleared.' 'Info'
        Set-Footer 'Grid cleared.'
    }
})

$btnClearLog.Add_Click({ $logBox.Clear() })
$btnFind.Add_Click({ Show-FindReplace })
$btnTypes.Add_Click({ Show-ContentTypePicker $btnTypes })

$form.Add_FormClosing({
    if ($script:Connected) { try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {} }
})
#endregion

# ─────────────────────────────────────────────────────────────────────────────
#region Find & Replace dialog
# ─────────────────────────────────────────────────────────────────────────────
# Edits the New Name / New Description columns in place. Supports literal or regex
# replace, case sensitivity, name/description targeting, all-vs-selected scope, and
# a one-click "strip trailing version" preset (e.g. removing " v3.9" / "-2.8").
function Show-FindReplace {
    if ($grid.Rows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('Pull or import profiles first.', 'Nothing to Edit', 'OK', 'Information') | Out-Null
        return
    }
    try { $grid.EndEdit() | Out-Null } catch {}

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text            = 'Find & Replace'
    $dlg.ClientSize      = New-Object System.Drawing.Size(484, 416)
    $dlg.StartPosition   = 'CenterParent'
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox     = $false
    $dlg.MinimizeBox     = $false
    $dlg.Font            = $FontUI
    $dlg.BackColor       = $Theme.White
    try { $dlg.Icon = $form.Icon } catch {}

    # Layout via TableLayoutPanel + docking (same approach as the main window) so the whole
    # dialog scales uniformly and never clips its controls on high-DPI displays.
    $rootT = New-Object System.Windows.Forms.TableLayoutPanel
    $rootT.Dock = 'Fill'; $rootT.ColumnCount = 1; $rootT.RowCount = 3
    $rootT.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
    $rootT.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 46)))  | Out-Null
    $rootT.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
    $rootT.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 58)))  | Out-Null
    $dlg.Controls.Add($rootT)

    # Gradient header
    $hd = New-Object System.Windows.Forms.Panel
    $hd.Dock = 'Fill'
    $hd.Add_Paint({
        param($s, $e)
        $r = $s.ClientRectangle
        if ($r.Width -le 0) { return }
        $b = New-Object System.Drawing.Drawing2D.LinearGradientBrush($r, $Theme.GradLeft, $Theme.GradRight, 0.0)
        $e.Graphics.FillRectangle($b, $r); $b.Dispose()
        $tb = New-Object System.Drawing.SolidBrush($Theme.White)
        $e.Graphics.DrawString('Find & Replace', (New-Object System.Drawing.Font('Segoe UI Light', 15)), $tb, 14, 9)
        $tb.Dispose()
    })
    $rootT.Controls.Add($hd, 0, 0)

    # Content — one column, auto-height rows; text boxes Dock=Fill to stretch full width
    $ct = New-Object System.Windows.Forms.TableLayoutPanel
    $ct.Dock = 'Fill'; $ct.ColumnCount = 1
    $ct.Padding = New-Object System.Windows.Forms.Padding(16, 8, 16, 4)
    $ct.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
    $rootT.Controls.Add($ct, 0, 1)

    function Add-CtRow { param($Control)
        $ct.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null
        $ct.Controls.Add($Control, 0, ($ct.RowStyles.Count - 1))
        $ct.RowCount = $ct.RowStyles.Count
    }
    function New-Lbl { param($Text, [bool]$Bold = $false)
        $l = New-Object System.Windows.Forms.Label
        $l.Text = $Text; $l.AutoSize = $true
        $l.Margin = New-Object System.Windows.Forms.Padding(0, 6, 0, 1)
        $l.Font = if ($Bold) { $FontUIBold } else { $FontUI }
        return $l
    }
    function New-RowFlow {
        $f = New-Object System.Windows.Forms.FlowLayoutPanel
        $f.AutoSize = $true; $f.AutoSizeMode = 'GrowAndShrink'; $f.WrapContents = $false
        $f.Dock = 'Fill'; $f.Margin = New-Object System.Windows.Forms.Padding(0, 4, 0, 0)
        return $f
    }

    Add-CtRow (New-Lbl 'Find:')
    $txtFind = New-Object System.Windows.Forms.TextBox
    $txtFind.Dock = 'Fill'; $txtFind.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 2)
    Add-CtRow $txtFind

    Add-CtRow (New-Lbl 'Replace with:  (leave blank to delete the matched text)')
    $txtRepl = New-Object System.Windows.Forms.TextBox
    $txtRepl.Dock = 'Fill'; $txtRepl.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 2)
    Add-CtRow $txtRepl

    $optsFlow = New-RowFlow
    $chkRegex = New-Object System.Windows.Forms.CheckBox
    $chkRegex.Text = 'Use regular expression'; $chkRegex.AutoSize = $true
    $chkRegex.Margin = New-Object System.Windows.Forms.Padding(0, 3, 28, 3)
    $chkCase = New-Object System.Windows.Forms.CheckBox
    $chkCase.Text = 'Case sensitive'; $chkCase.AutoSize = $true
    $chkCase.Margin = New-Object System.Windows.Forms.Padding(0, 3, 0, 3)
    $optsFlow.Controls.AddRange(@($chkRegex, $chkCase))
    Add-CtRow $optsFlow

    $applyFlow = New-RowFlow
    $lblApply = New-Lbl 'Apply to:' $true; $lblApply.Margin = New-Object System.Windows.Forms.Padding(0, 4, 12, 0)
    $chkName = New-Object System.Windows.Forms.CheckBox
    $chkName.Text = 'New Name'; $chkName.AutoSize = $true; $chkName.Checked = $true
    $chkName.Margin = New-Object System.Windows.Forms.Padding(0, 2, 18, 0)
    $chkDesc = New-Object System.Windows.Forms.CheckBox
    $chkDesc.Text = 'New Description'; $chkDesc.AutoSize = $true
    $chkDesc.Margin = New-Object System.Windows.Forms.Padding(0, 2, 0, 0)
    $applyFlow.Controls.AddRange(@($lblApply, $chkName, $chkDesc))
    Add-CtRow $applyFlow

    $scopeFlow = New-RowFlow
    $lblScope = New-Lbl 'Scope:' $true; $lblScope.Margin = New-Object System.Windows.Forms.Padding(0, 4, 23, 0)
    $rdoAll = New-Object System.Windows.Forms.RadioButton
    $rdoAll.Text = 'All rows'; $rdoAll.AutoSize = $true; $rdoAll.Checked = $true
    $rdoAll.Margin = New-Object System.Windows.Forms.Padding(0, 2, 18, 0)
    $rdoSel = New-Object System.Windows.Forms.RadioButton
    $rdoSel.Text = 'Selected rows only'; $rdoSel.AutoSize = $true
    $rdoSel.Margin = New-Object System.Windows.Forms.Padding(0, 2, 0, 0)
    $scopeFlow.Controls.AddRange(@($lblScope, $rdoAll, $rdoSel))
    Add-CtRow $scopeFlow

    $btnPreset = New-Object System.Windows.Forms.Button
    $btnPreset.Text = 'Strip trailing version (e.g. v3.9, -2.8)'
    Set-ToolButton -Btn $btnPreset
    $btnPreset.Margin = New-Object System.Windows.Forms.Padding(0, 10, 0, 2)
    Add-CtRow $btnPreset
    $btnPreset.Add_Click({
        $chkRegex.Checked = $true
        $txtFind.Text = '[\s_\-]*[vV]?\d+(\.\d+)+\s*$'
        $txtRepl.Text = ''
        $chkName.Checked = $true
    })

    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.AutoSize = $true
    $lblStatus.Margin = New-Object System.Windows.Forms.Padding(0, 8, 0, 0)
    $lblStatus.ForeColor = C '#555555'
    Add-CtRow $lblStatus

    # Footer — action buttons docked to the right so they are always visible
    $foot = New-Object System.Windows.Forms.Panel
    $foot.Dock = 'Fill'
    $footFlow = New-Object System.Windows.Forms.FlowLayoutPanel
    $footFlow.Dock = 'Right'; $footFlow.FlowDirection = 'LeftToRight'; $footFlow.WrapContents = $false
    $footFlow.AutoSize = $true; $footFlow.AutoSizeMode = 'GrowAndShrink'
    $footFlow.Padding = New-Object System.Windows.Forms.Padding(0, 14, 12, 0)
    $foot.Controls.Add($footFlow)
    $rootT.Controls.Add($foot, 0, 2)

    $btnPreview = New-Object System.Windows.Forms.Button; $btnPreview.Text = 'Preview'
    $btnDoRepl  = New-Object System.Windows.Forms.Button; $btnDoRepl.Text  = 'Replace'
    $btnClose   = New-Object System.Windows.Forms.Button; $btnClose.Text   = 'Close'
    Set-ToolButton    -Btn $btnPreview
    Set-PrimaryButton -Btn $btnDoRepl -Back $Theme.Primary
    Set-ToolButton    -Btn $btnClose
    foreach ($b in @($btnPreview, $btnDoRepl, $btnClose)) {
        $b.MinimumSize = New-Object System.Drawing.Size(80, 30)
        $b.Margin      = New-Object System.Windows.Forms.Padding(0, 0, 7, 0)
    }
    $footFlow.Controls.AddRange(@($btnPreview, $btnDoRepl, $btnClose))

    # Core replace routine — $PreviewOnly counts without modifying the grid
    $run = {
        param([bool]$PreviewOnly)
        $find = $txtFind.Text
        if ([string]::IsNullOrEmpty($find)) {
            $lblStatus.ForeColor = $Theme.LogFail; $lblStatus.Text = 'Enter something to find.'; return
        }
        if (-not $chkName.Checked -and -not $chkDesc.Checked) {
            $lblStatus.ForeColor = $Theme.LogFail; $lblStatus.Text = 'Choose at least one target (Name / Description).'; return
        }
        $cols = @()
        if ($chkName.Checked) { $cols += 'NewName' }
        if ($chkDesc.Checked) { $cols += 'NewDescription' }

        $opts = if ($chkCase.Checked) { [System.Text.RegularExpressions.RegexOptions]::None }
                else { [System.Text.RegularExpressions.RegexOptions]::IgnoreCase }
        $pattern = if ($chkRegex.Checked) { $find } else { [System.Text.RegularExpressions.Regex]::Escape($find) }
        try { $re = [System.Text.RegularExpressions.Regex]::new($pattern, $opts) }
        catch { $lblStatus.ForeColor = $Theme.LogFail; $lblStatus.Text = "Invalid regex: $($_.Exception.Message)"; return }

        $replText = $txtRepl.Text
        # Literal mode: route replacement through an evaluator so $ and \ aren't treated as substitutions
        $evaluator = [System.Text.RegularExpressions.MatchEvaluator] { param($m) $replText }.GetNewClosure()

        $rows = if ($rdoSel.Checked) { @($grid.SelectedRows) } else { @($grid.Rows) }
        if ($rdoSel.Checked -and $rows.Count -eq 0) {
            $lblStatus.ForeColor = $Theme.LogFail; $lblStatus.Text = 'No rows are selected.'; return
        }

        $cells = 0; $rowsHit = @{}
        foreach ($r in $rows) {
            if ($r.IsNewRow) { continue }
            foreach ($col in $cols) {
                $old = [string]$r.Cells[$col].Value
                $new = if ($chkRegex.Checked) { $re.Replace($old, $replText) } else { $re.Replace($old, $evaluator) }
                if ($new -ne $old) {
                    $cells++; $rowsHit[$r.Index] = $true
                    if (-not $PreviewOnly) { $r.Cells[$col].Value = $new }
                }
            }
        }

        if ($PreviewOnly) {
            $lblStatus.ForeColor = C '#555555'
            $lblStatus.Text = "Preview: $cells cell(s) in $($rowsHit.Count) row(s) would change."
        } else {
            Update-AllHighlights
            $lblStatus.ForeColor = $Theme.LogOk
            $lblStatus.Text = "Replaced in $cells cell(s) across $($rowsHit.Count) row(s)."
            if ($cells -gt 0) {
                Write-Log "Find & Replace: '$find' -> '$replText' changed $cells cell(s) in $($rowsHit.Count) row(s)$(if ($chkRegex.Checked) {' [regex]'})." 'OK'
            }
        }
    }.GetNewClosure()

    $btnPreview.Add_Click({ & $run $true })
    $btnDoRepl.Add_Click({  & $run $false })
    $btnClose.Add_Click({ $dlg.Close() })

    $dlg.AcceptButton = $btnDoRepl
    $dlg.CancelButton = $btnClose
    [void]$dlg.ShowDialog($form)
    $dlg.Dispose()
}
#endregion

# ─────────────────────────────────────────────────────────────────────────────
#region Content Types picker
# ─────────────────────────────────────────────────────────────────────────────
# A dropdown-style popup (checklist + Select all) under the Content Types button.
# The ticked set drives what the next Pull fetches. Closes when it loses focus.
function Show-ContentTypePicker {
    param($AnchorButton)

    $pop = New-Object System.Windows.Forms.Form
    $pop.FormBorderStyle = 'None'
    $pop.StartPosition   = 'Manual'
    $pop.ShowInTaskbar   = $false
    $pop.KeyPreview      = $true
    $pop.BackColor       = C '#B9B9B9'                              # shows as a 1px border
    $pop.Padding         = New-Object System.Windows.Forms.Padding(1)
    $pop.Size            = New-Object System.Drawing.Size(268, 474)

    $inner = New-Object System.Windows.Forms.Panel
    $inner.Dock = 'Fill'; $inner.BackColor = $Theme.White
    $inner.Padding = New-Object System.Windows.Forms.Padding(12, 10, 12, 10)
    $pop.Controls.Add($inner)

    $clb = New-Object System.Windows.Forms.CheckedListBox
    $clb.Dock = 'Fill'; $clb.Font = $FontUI; $clb.CheckOnClick = $true
    $clb.BorderStyle = 'None'; $clb.IntegralHeight = $false
    foreach ($k in $ContentTypes.Keys) { [void]$clb.Items.Add($k) }

    $sep = New-Object System.Windows.Forms.Panel
    $sep.Dock = 'Top'; $sep.Height = 8; $sep.BackColor = $Theme.White

    $chkAll = New-Object System.Windows.Forms.CheckBox
    $chkAll.Text = 'Select all'; $chkAll.Dock = 'Top'; $chkAll.Height = 26; $chkAll.Font = $FontUI

    $hdr = New-Object System.Windows.Forms.Label
    $hdr.Text = 'Content Types'; $hdr.Dock = 'Top'; $hdr.Height = 24; $hdr.Font = $FontUIBold

    # Fill added first (sits behind), then Top items in bottom-to-top add order
    $inner.Controls.Add($clb)
    $inner.Controls.Add($sep)
    $inner.Controls.Add($chkAll)
    $inner.Controls.Add($hdr)

    for ($i = 0; $i -lt $clb.Items.Count; $i++) {
        $clb.SetItemChecked($i, ($global:IPMSelectedContentTypes -contains ([string]$clb.Items[$i])))
    }
    $chkAll.Checked = ($clb.CheckedItems.Count -eq $clb.Items.Count)

    $chkAll.Add_Click({
        $state = $chkAll.Checked
        for ($i = 0; $i -lt $clb.Items.Count; $i++) { $clb.SetItemChecked($i, $state) }
    }.GetNewClosure())

    $pop.Add_KeyDown({ param($s, $e) if ($e.KeyCode -eq 'Escape') { $pop.Close() } }.GetNewClosure())
    $pop.Add_Deactivate({ $pop.Close() }.GetNewClosure())
    $pop.Add_FormClosed({
        $sel = [System.Collections.Generic.List[string]]::new()
        foreach ($it in $clb.CheckedItems) { $sel.Add([string]$it) }
        $global:IPMSelectedContentTypes = $sel
        Set-Footer "$($sel.Count) of $($ContentTypes.Count) content types selected for the next Pull."
    }.GetNewClosure())

    # Position under the button, clamped to the screen working area
    $pt = $AnchorButton.PointToScreen([System.Drawing.Point]::new(0, $AnchorButton.Height + 2))
    $wa = [System.Windows.Forms.Screen]::FromControl($AnchorButton).WorkingArea
    $x  = [Math]::Min([int]$pt.X, $wa.Right - $pop.Width - 4)
    $y  = [int]$pt.Y
    if ($y + $pop.Height -gt $wa.Bottom) { $y = $wa.Bottom - $pop.Height - 4 }
    $pop.Location = New-Object System.Drawing.Point($x, $y)

    $pop.Show($form)
    $pop.Activate()
}
#endregion

# ─────────────────────────────────────────────────────────────────────────────
#region Launch
# ─────────────────────────────────────────────────────────────────────────────
Write-Log 'Intune Profile Manager started.' 'Info'
Write-Log 'Connect, choose Content Types if needed, then Pull.' 'Info'
[void]$form.ShowDialog()
#endregion
