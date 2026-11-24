<#
.SYNOPSIS
    IRCleanup - Intune Remediation Cleanup.
    Search, back up and bulk-delete Intune Remediation scripts (deviceHealthScripts) from a WPF grid.

.DESCRIPTION
    Talks to Microsoft Graph via Invoke-MgGraphRequest against
    beta/deviceManagement/deviceHealthScripts, following @odata.nextLink paging.

    Features:
      - Search (every term must match name, description, publisher or ID), sort and filter
        (hide Microsoft scripts, only unassigned, only checked). Export the view to CSV.
      - Bulk delete of checked scripts. Each one is backed up first; if its backup fails it is
        not deleted. Microsoft (global) scripts can't be checked or deleted.
      - "Back up checked" and "Back up all" back up scripts without deleting anything.

    Backup layout, in the backup folder:
      <script name>\script.json      Raw Graph object incl. assignments (restorable as-is)
      <script name>\settings.txt     The portal's Settings section, human-readable
      <script name>\assignments.txt  Included/excluded groups by name, with schedule and filter
      <script name>\detect.ps1       Detection script
      <script name>\remediate.ps1    Remediation script, if the script has one
      BackupLog_<timestamp>.csv      One per backup run, with each script's description
      DeleteLog_<timestamp>.csv      One per delete run, with each script's description

    A different script with the same name goes to "<script name> (2)"; backing up the same script
    again overwrites its folder. The backup folder defaults to %APPDATA%\IRCleanup\Backups and can
    be changed in the window (remembered for next time) or with -BackupPath.

.PARAMETER ClientId
    App registration client ID. If omitted, the Microsoft Graph PowerShell app is used.

.PARAMETER TenantId
    Tenant ID. If omitted, you sign in to your account's home tenant.

.PARAMETER BackupPath
    Backup folder for this run (local path or UNC share). Used by the backup buttons and the
    pre-delete backup. Overrides the folder remembered from the window, without changing it.

.PARAMETER NoBackup
    Skip the pre-delete backup. Not recommended. The backup buttons still work.

.PARAMETER Theme
    Dark or Light. Defaults to Dark.

.EXAMPLE
    .\IRCleanup.ps1
    Opens the grid, backing up to the remembered folder (or %APPDATA%\IRCleanup\Backups).

.EXAMPLE
    .\IRCleanup.ps1 -BackupPath '\\fileserver\IT\IntuneBackups' -Theme Light
    Backs up to a network share for this run, using the light theme.

.NOTES
    Delegated permissions:
      DeviceManagementConfiguration.ReadWrite.All  Required - list, back up and delete scripts
      GroupMember.Read.All                         Optional - group names in assignments.txt.
                                                   Without it, sign-in is retried without it and
                                                   group IDs are shown instead.

    If an Intune Multi Admin Approval access policy covers scripts, deletes fail with
    "Header 'x-msft-approval-justification' is required". The backup is still taken.
#>
[CmdletBinding()]
param(
    [string]$ClientId,
    [string]$TenantId,
    [string]$BackupPath,
    [switch]$NoBackup,
    [ValidateSet('Dark', 'Light')]
    [string]$Theme
)

$ErrorActionPreference = 'Stop'

# --- WPF needs an STA thread -------------------------------------------------
if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    $exe = (Get-Process -Id $PID).Path
    $relaunch = @('-NoProfile', '-STA', '-File', $PSCommandPath) + ($PSBoundParameters.GetEnumerator() | ForEach-Object {
        if ($_.Value -is [switch]) { if ($_.Value) { "-$($_.Key)" } } else { "-$($_.Key)"; "$($_.Value)" }
    })
    & $exe @relaunch
    return
}

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

$script:GraphBaseUrl = 'https://graph.microsoft.com/beta'
$script:DefaultBackupRoot = Join-Path $env:APPDATA 'IRCleanup\Backups'
$script:SettingsPath      = Join-Path $env:APPDATA 'IRCleanup\settings.json'
# PS 7 omits type info by default and warns that -NoTypeInformation is obsolete; 5.1 still needs it
$script:CsvArgs = @{ Encoding = 'UTF8' }
if ($PSVersionTable.PSVersion.Major -lt 6) { $script:CsvArgs.NoTypeInformation = $true }

# --- Row model (INotifyPropertyChanged so checkboxes/status update live) -----
if (-not ('IRCleanup.RemediationRow' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.ComponentModel;
namespace IRCleanup {
    public class RemediationRow : INotifyPropertyChanged {
        private bool _isSelected;
        private string _status = "";
        public event PropertyChangedEventHandler PropertyChanged;
        private void Notify(string name) {
            PropertyChangedEventHandler h = PropertyChanged;
            if (h != null) h(this, new PropertyChangedEventArgs(name));
        }
        public bool IsSelected {
            get { return _isSelected; }
            set { if (_isSelected != value && (CanDelete || !value)) { _isSelected = value; Notify("IsSelected"); } }
        }
        public string Status {
            get { return _status; }
            set { if (_status != value) { _status = value; Notify("Status"); } }
        }
        public string Id { get; set; }
        public string DisplayName { get; set; }
        public string Description { get; set; }
        public string Publisher { get; set; }
        public string Version { get; set; }
        public bool IsGlobal { get; set; }
        public bool CanDelete { get { return !IsGlobal; } }
        public string Source { get { return IsGlobal ? "Microsoft" : "Custom"; } }
        public int AssignmentCount { get; set; }   // -1 = unknown
        public string Assigned {
            get { return AssignmentCount < 0 ? "?" : (AssignmentCount == 0 ? "No" : "Yes (" + AssignmentCount + ")"); }
        }
        public DateTime? Created { get; set; }
        public DateTime? Modified { get; set; }
    }
}
'@
}

# --- Theme -------------------------------------------------------------------
if (-not $Theme) { $Theme = 'Dark' }
$t = if ($Theme -eq 'Light') {
    @{ Window = '#F5F5F5'; Control = '#FFFFFF'; Alt = '#F7F7F7'; Hover = '#E5E5E5'; Text = '#1E1E1E'; Text2 = '#666666'
       Placeholder = '#999999'; Border = '#CCCCCC'; Header = '#E8E8E8'; Selected = '#0078D4'; SelectedText = '#FFFFFF'
       Button = '#E1E1E1'; ButtonHover = '#D5D5D5'; Green = '#2E8B57'; GreenHover = '#3E9B67'; Red = '#C42B1C'; RedHover = '#D43B2C'; Blue = '#0078D4' }
} else {
    @{ Window = '#1E1E1E'; Control = '#2D2D30'; Alt = '#2A2A2D'; Hover = '#3E3E42'; Text = '#FFFFFF'; Text2 = '#CCCCCC'
       Placeholder = '#808080'; Border = '#3F3F46'; Header = '#2D2D30'; Selected = '#094771'; SelectedText = '#FFFFFF'
       Button = '#2D2D30'; ButtonHover = '#3E3E42'; Green = '#2E8B57'; GreenHover = '#3E9B67'; Red = '#B3261E'; RedHover = '#D43B2C'; Blue = '#0E639C' }
}

# --- Graph helpers -----------------------------------------------------------
function Install-RequiredModules {
    if (-not (Get-Module -Name Microsoft.Graph.Authentication -ListAvailable)) {
        Write-Host "Installing Microsoft.Graph.Authentication..." -ForegroundColor Yellow
        Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module Microsoft.Graph.Authentication
}

function Connect-ToGraph {
    param([string]$ClientId, [string]$TenantId)
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Green
    $p = @{ NoWelcome = $true; ErrorAction = 'Stop'; WarningAction = 'SilentlyContinue' }
    if ($ClientId) { $p.ClientId = $ClientId }
    if ($TenantId) { $p.TenantId = $TenantId }
    try {
        Connect-MgGraph @p -Scopes 'DeviceManagementConfiguration.ReadWrite.All', 'GroupMember.Read.All' | Out-Null
    }
    catch {
        # GroupMember.Read.All only turns group IDs into names in assignments.txt. If it can't be
        # granted (e.g. no admin consent), carry on without it rather than blocking the tool.
        Write-Host "Sign-in with GroupMember.Read.All failed: $($_.Exception.Message)" -ForegroundColor Yellow
        Write-Host "Retrying without it - assignments.txt will show group IDs instead of names." -ForegroundColor Yellow
        try { Connect-MgGraph @p -Scopes 'DeviceManagementConfiguration.ReadWrite.All' | Out-Null }
        catch { Write-Host "Auth failed: $_" -ForegroundColor Red; return $false }
    }
    $ctx = Get-MgContext
    if ($ctx.Scopes -notcontains 'DeviceManagementConfiguration.ReadWrite.All') {
        Write-Host "Warning: token lacks DeviceManagementConfiguration.ReadWrite.All - deletes will fail." -ForegroundColor Yellow
    }
    if (-not ($ctx.Scopes | Where-Object { $_ -in 'GroupMember.Read.All', 'Group.Read.All', 'Group.ReadWrite.All', 'Directory.Read.All', 'Directory.ReadWrite.All' })) {
        Write-Host "Note: no group read permission - assignments.txt will show group IDs instead of names." -ForegroundColor Yellow
    }
    Write-Host "Connected as: $($ctx.Account)" -ForegroundColor Green
    $true
}

function Get-GraphErrorText {
    # Invoke-MgGraphRequest hides Graph's explanation; it lives in ErrorDetails as JSON
    param($ErrorRecord)
    $text = $ErrorRecord.Exception.Message
    $details = "$($ErrorRecord.ErrorDetails.Message)"
    if ($details) {
        try {
            $e = ($details | ConvertFrom-Json).error
            $inner = "$($e.message)"
            # Intune often nests a second JSON payload inside error.message
            if ($inner -match '"Message"\s*:\s*"([^"]+)') { $inner = $Matches[1] }
            $text = "$text | $($e.code): $inner"
        }
        catch { $text = "$text | $details" }
    }
    $text
}

function Invoke-Graph {
    param([string]$Uri, [string]$Method = 'GET')
    Invoke-MgGraphRequest -Method $Method -Uri $Uri
}

function Get-AllPages {
    param([string]$Uri)
    $resp = Invoke-Graph -Uri $Uri
    $items = @($resp.value)
    while ($resp.'@odata.nextLink') {
        $resp = Invoke-Graph -Uri $resp.'@odata.nextLink'
        $items += $resp.value
    }
    $items
}

function Get-RemediationRows {
    Write-Host "Fetching remediation scripts..." -ForegroundColor White
    $base = "$script:GraphBaseUrl/deviceManagement/deviceHealthScripts"
    $haveAssignments = $true
    try   { $scripts = Get-AllPages -Uri "$base`?`$expand=assignments" }
    catch {
        Write-Host "Could not expand assignments, loading without them." -ForegroundColor Yellow
        $haveAssignments = $false
        $scripts = Get-AllPages -Uri $base
    }

    foreach ($s in $scripts) {
        $row = [IRCleanup.RemediationRow]::new()
        $row.Id          = $s.id
        $row.DisplayName = $s.displayName
        $row.Description = ("$($s.description)" -replace '\s+', ' ').Trim()
        $row.Publisher   = $s.publisher
        $row.Version     = $s.version
        $row.IsGlobal    = [bool]$s.isGlobalScript
        $row.AssignmentCount = if ($haveAssignments) { @($s.assignments).Where({ $_ }).Count } else { -1 }
        if ($s.createdDateTime)      { $row.Created  = ([datetime]$s.createdDateTime).ToLocalTime() }
        if ($s.lastModifiedDateTime) { $row.Modified = ([datetime]$s.lastModifiedDateTime).ToLocalTime() }
        $row
    }
}

function Get-BackupFolderFor {
    # <root>\<script name>. Display names aren't unique, so a different script with the same name
    # gets "<name> (2)", "(3)"... A folder already holding THIS script's backup is reused, and any
    # folder without a matching script.json (e.g. unrelated folders in the root) is never touched.
    param([string]$Root, [string]$Name, [string]$Id)
    $safe = ($Name -replace '[\\/:*?"<>|]', '_').Trim().TrimEnd('.', ' ')
    if (-not $safe) { $safe = $Id }
    for ($n = 1; ; $n++) {
        $dir = Join-Path $Root $(if ($n -eq 1) { $safe } else { "$safe ($n)" })
        if (-not [IO.Directory]::Exists($dir)) { return $dir }
        try {
            $existing = [IO.File]::ReadAllText((Join-Path $dir 'script.json')) | ConvertFrom-Json
            if ($existing.id -eq $Id) { return $dir }
        } catch { }
    }
}

function Get-PortalSettings {
    # The Intune portal's "Settings" section, with its labels and Yes/No values.
    # Graph names differ and two are inverted: runAsAccount 'user' = logged-on credentials, runAs32Bit false = 64-bit.
    param($Script)
    $yn = { param($v) if ($v) { 'Yes' } else { 'No' } }
    [ordered]@{
        'Detection script'                                = & $yn $Script.detectionScriptContent
        'Remediation script'                              = & $yn $Script.remediationScriptContent
        'Run this script using the logged-on credentials' = & $yn ($Script.runAsAccount -eq 'user')
        'Enforce script signature check'                  = & $yn $Script.enforceSignatureCheck
        'Run script in 64-bit PowerShell'                 = & $yn (-not $Script.runAs32Bit)
    }
}

function Format-PortalSettings {
    # Human-readable page laid out like the portal's Settings section (CRLF so Notepad is happy)
    param($Script)
    $settings = Get-PortalSettings $Script
    $width = ($settings.Keys | ForEach-Object Length | Measure-Object -Maximum).Maximum + 4
    $name  = "$($Script.displayName)"
    $lines = @($name, ('=' * $name.Length), '', 'Settings', '--------')
    $lines += foreach ($k in $settings.Keys) { $k.PadRight($width) + $settings[$k] }
    ($lines -join "`r`n") + "`r`n"
}

# Group and filter names, looked up once per session
$script:GroupNames  = @{}
$script:FilterNames = @{}

function Get-GroupName {
    # Never throws: a missing name shouldn't fail a backup (script.json still has the ID)
    param([string]$Id)
    if (-not $script:GroupNames.ContainsKey($Id)) {
        $script:GroupNames[$Id] = try {
            "$((Invoke-Graph -Uri "https://graph.microsoft.com/v1.0/groups/$Id`?`$select=displayName").displayName)"
        }
        catch {
            $err = Get-GraphErrorText $_
            if ($err -match 'NotFound|Request_ResourceNotFound') { '(group not found - deleted?)' }
            elseif ($err -match 'Forbidden|Authorization_RequestDenied|Insufficient privileges') { '(name unavailable - needs GroupMember.Read.All)' }
            else { '(name unavailable)' }
        }
    }
    $script:GroupNames[$Id]
}

function Get-FilterText {
    # "None", or e.g. "Include - Windows 11 only"
    param($Target)
    $mode = "$($Target.deviceAndAppManagementAssignmentFilterType)"
    $id   = "$($Target.deviceAndAppManagementAssignmentFilterId)"
    if (-not $mode -or $mode -eq 'none' -or -not $id -or $id -eq [guid]::Empty.ToString()) { return 'None' }
    if (-not $script:FilterNames.ContainsKey($id)) {
        $script:FilterNames[$id] = try { "$((Invoke-Graph -Uri "$script:GraphBaseUrl/deviceManagement/assignmentFilters/$id").displayName)" }
                                   catch { "(filter $id)" }
    }
    "$($mode.Substring(0, 1).ToUpper())$($mode.Substring(1)) - $($script:FilterNames[$id])"
}

function Format-RunSchedule {
    param($Schedule)
    if (-not $Schedule) { return 'Not set' }
    $zone = if ($Schedule.useUtc) { 'UTC' } else { 'device local time' }
    $time = if ("$($Schedule.time)" -match '^\d{1,2}:\d{2}') { $Matches[0] } else { "$($Schedule.time)" }
    $date = if ($Schedule.date -is [datetime]) { $Schedule.date.ToString('yyyy-MM-dd') } else { "$($Schedule.date)" }
    $n    = [int]$Schedule.interval
    switch -Wildcard ("$($Schedule.'@odata.type')") {
        '*RunOnceSchedule' { "Once on $date at $time ($zone)" }
        '*HourlySchedule'  { if ($n -le 1) { 'Every hour' } else { "Every $n hours" } }
        '*DailySchedule'   { "$(if ($n -le 1) { 'Every day' } else { "Every $n days" }) at $time ($zone)" }
        default            { $Schedule | ConvertTo-Json -Compress }
    }
}

function Format-Assignments {
    # Human-readable list of who the script is assigned to, with group names instead of IDs
    param($Script)
    $name  = "$($Script.displayName)"
    $lines = [System.Collections.Generic.List[string]]::new()
    foreach ($l in @($name, ('=' * $name.Length), '', 'Assignments', '-----------')) { $lines.Add($l) }

    $all = @($Script.assignments | Where-Object { $_ })
    if (-not $all) { $lines.Add('Not assigned.') }
    else {
        $excluded = @($all | Where-Object { "$($_.target.'@odata.type')" -like '*exclusionGroupAssignmentTarget' })
        $included = @($all | Where-Object { "$($_.target.'@odata.type')" -notlike '*exclusionGroupAssignmentTarget' })

        $lines.Add('Included')
        if (-not $included) { $lines.Add('  None') }
        foreach ($a in $included) {
            $t = $a.target
            switch -Wildcard ("$($t.'@odata.type')") {
                '*allDevicesAssignmentTarget'       { $lines.Add('  All devices') }
                '*allLicensedUsersAssignmentTarget' { $lines.Add('  All users') }
                '*groupAssignmentTarget'            { $lines.Add("  $(Get-GroupName $t.groupId)"); $lines.Add("    Group ID:  $($t.groupId)") }
                default                             { $lines.Add("  $($t.'@odata.type' -replace '^#microsoft\.graph\.', '')") }
            }
            $lines.Add("    Schedule:  $(Format-RunSchedule $a.runSchedule)")
            $lines.Add("    Filter:    $(Get-FilterText $t)")
        }

        $lines.Add('')
        $lines.Add('Excluded')
        if (-not $excluded) { $lines.Add('  None') }
        foreach ($a in $excluded) {
            $lines.Add("  $(Get-GroupName $a.target.groupId)")
            $lines.Add("    Group ID:  $($a.target.groupId)")
        }
    }
    ($lines -join "`r`n") + "`r`n"
}

function Save-RemediationBackup {
    param([string]$Id, [string]$Root)
    $full = Invoke-Graph -Uri "$script:GraphBaseUrl/deviceManagement/deviceHealthScripts/$Id`?`$expand=assignments"
    $dir  = Get-BackupFolderFor -Root $Root -Name $full.displayName -Id $Id
    # .NET IO throughout: script names can contain [ ], which PowerShell cmdlets treat as wildcards
    [IO.Directory]::CreateDirectory($dir) | Out-Null
    $utf8 = [Text.UTF8Encoding]::new($false)
    # script.json = raw Graph object (restorable as-is); settings.txt = the same settings as the portal shows them
    [IO.File]::WriteAllText((Join-Path $dir 'script.json'), ($full | ConvertTo-Json -Depth 20), $utf8)
    [IO.File]::WriteAllText((Join-Path $dir 'settings.txt'), (Format-PortalSettings $full), $utf8)
    # Convenience file only - if it can't be built, say so in it rather than failing the backup (and blocking a delete)
    $assignText = try { Format-Assignments $full } catch { "Couldn't build the assignment summary: $($_.Exception.Message)`r`nThe raw assignments are in script.json.`r`n" }
    [IO.File]::WriteAllText((Join-Path $dir 'assignments.txt'), $assignText, $utf8)
    foreach ($pair in @(@('detectionScriptContent', 'detect.ps1'), @('remediationScriptContent', 'remediate.ps1'))) {
        $file = Join-Path $dir $pair[1]
        if ($full[$pair[0]]) { [IO.File]::WriteAllBytes($file, [Convert]::FromBase64String($full[$pair[0]])) }
        elseif ([IO.File]::Exists($file)) { [IO.File]::Delete($file) }   # don't leave a stale copy from an earlier backup
    }
    $dir
}

# --- Backup folder -----------------------------------------------------------
function Get-SavedBackupRoot {
    try {
        if (Test-Path -LiteralPath $script:SettingsPath) {
            return "$((Get-Content -LiteralPath $script:SettingsPath -Raw | ConvertFrom-Json).BackupRoot)"
        }
    } catch { }
    ''
}

function Save-BackupRoot {
    param([string]$Path)
    try {
        [IO.Directory]::CreateDirectory((Split-Path $script:SettingsPath)) | Out-Null
        [IO.File]::WriteAllText($script:SettingsPath, (@{ BackupRoot = $Path } | ConvertTo-Json), [Text.UTF8Encoding]::new($false))
    } catch { }
}

function Resolve-BackupRoot {
    # Normalises what's in the box; throws a user-facing message if it isn't usable
    param([string]$Text)
    $p = [Environment]::ExpandEnvironmentVariables($Text.Trim().Trim('"'))
    if (-not $p) { throw 'Choose a backup folder first.' }
    if (-not [IO.Path]::IsPathRooted($p)) { throw "Backup folder must be a full path, e.g. D:\Backups or \\server\share\Backups.`n`nGot: $p" }
    [IO.Path]::GetFullPath($p)
}

function Select-BackupFolder {
    param([string]$Start, $Owner)
    if ('Microsoft.Win32.OpenFolderDialog' -as [type]) {
        # .NET 8+ (PowerShell 7.4+): native WPF folder picker
        $dlg = [Microsoft.Win32.OpenFolderDialog]::new()
        $dlg.Title = 'Choose backup folder'
        if ($Start -and (Test-Path -LiteralPath $Start)) { $dlg.InitialDirectory = $Start }
        if ($dlg.ShowDialog($Owner)) { return $dlg.FolderName }
        return $null
    }
    # Windows PowerShell 5.1 / older .NET
    Add-Type -AssemblyName System.Windows.Forms
    $dlg = [System.Windows.Forms.FolderBrowserDialog]::new()
    $dlg.Description = 'Choose backup folder'
    $dlg.ShowNewFolderButton = $true
    if ($Start -and (Test-Path -LiteralPath $Start)) { $dlg.SelectedPath = $Start }
    if ($dlg.ShowDialog() -eq 'OK') { return $dlg.SelectedPath }
    $null
}

# --- Start -------------------------------------------------------------------
# -BackupPath > folder remembered from the window > default
$script:BackupRoot = if (-not [string]::IsNullOrWhiteSpace($BackupPath)) { $BackupPath }
                     elseif ($saved = Get-SavedBackupRoot) { $saved }
                     else { $script:DefaultBackupRoot }

Install-RequiredModules
if (-not (Connect-ToGraph -ClientId $ClientId -TenantId $TenantId)) { return }

$script:Rows = [System.Collections.ObjectModel.ObservableCollection[object]]::new()
foreach ($r in Get-RemediationRows) { $script:Rows.Add($r) }
Write-Host "Loaded $($script:Rows.Count) remediation(s)." -ForegroundColor Green

# --- UI ----------------------------------------------------------------------
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="IRCleanup - Intune Remediation Cleanup" Height="750" Width="1150"
        WindowStartupLocation="CenterScreen" Background="$($t.Window)">
    <Window.Resources>
        <!-- Flat template for the coloured buttons: the stock one paints its own hover/disabled
             colours over Background. Colours go in style setters, not attributes, because local
             values would override the triggers. -->
        <ControlTemplate x:Key="FlatButton" TargetType="Button">
            <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="1" Padding="{TemplateBinding Padding}" SnapsToDevicePixels="True">
                <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
        </ControlTemplate>
        <Style TargetType="Button">
            <Setter Property="Background" Value="$($t.Button)"/>
            <Setter Property="Foreground" Value="$($t.Text)"/>
            <Setter Property="BorderBrush" Value="$($t.Border)"/>
            <Setter Property="Padding" Value="12,6"/>
            <Setter Property="Margin" Value="0,0,8,0"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Cursor" Value="Hand"/>
        </Style>
        <!-- Backup buttons: green, lighter on hover, faded when disabled -->
        <Style x:Key="GreenButton" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="Background" Value="$($t.Green)"/>
            <Setter Property="BorderBrush" Value="$($t.GreenHover)"/>
            <Setter Property="Template" Value="{StaticResource FlatButton}"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="$($t.GreenHover)"/>
                </Trigger>
                <Trigger Property="IsEnabled" Value="False">
                    <Setter Property="Opacity" Value="0.5"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style TargetType="CheckBox">
            <Setter Property="Foreground" Value="$($t.Text2)"/>
            <Setter Property="VerticalAlignment" Value="Center"/>
            <Setter Property="Margin" Value="0,0,16,0"/>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Background" Value="$($t.Control)"/>
            <Setter Property="Foreground" Value="$($t.Text)"/>
            <Setter Property="BorderBrush" Value="$($t.Border)"/>
            <Setter Property="CaretBrush" Value="$($t.Text)"/>
        </Style>
    </Window.Resources>
    <Grid Margin="15">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <Grid Grid.Row="0" Margin="0,0,0,10">
            <TextBox Name="SearchBox" Padding="8" FontSize="14"/>
            <TextBlock Name="SearchPlaceholder" Margin="10,0,0,0" VerticalAlignment="Center" IsHitTestVisible="False"
                       Foreground="$($t.Placeholder)" FontSize="14" Text="Search name, description, publisher or ID..."/>
        </Grid>

        <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,0,0,10">
            <CheckBox Name="HideGlobal" Content="Hide Microsoft scripts" IsChecked="True"/>
            <CheckBox Name="OnlyUnassigned" Content="Only unassigned"/>
            <CheckBox Name="OnlySelected" Content="Only checked"/>
            <TextBlock Name="CountText" Foreground="$($t.Text2)" VerticalAlignment="Center" Margin="10,0,0,0"/>
        </StackPanel>

        <DataGrid Name="Grid" Grid.Row="2" AutoGenerateColumns="False" CanUserAddRows="False" CanUserDeleteRows="False"
                  CanUserSortColumns="True" CanUserResizeColumns="True" SelectionMode="Extended" SelectionUnit="FullRow"
                  HeadersVisibility="Column" GridLinesVisibility="Horizontal" IsReadOnly="False"
                  Background="$($t.Control)" RowBackground="$($t.Control)" AlternatingRowBackground="$($t.Alt)"
                  Foreground="$($t.Text)" BorderBrush="$($t.Border)" HorizontalGridLinesBrush="$($t.Border)"
                  EnableRowVirtualization="True" FontSize="13">
            <DataGrid.Resources>
                <Style TargetType="DataGridColumnHeader">
                    <Setter Property="Background" Value="$($t.Header)"/>
                    <Setter Property="Foreground" Value="$($t.Text)"/>
                    <Setter Property="BorderBrush" Value="$($t.Border)"/>
                    <Setter Property="BorderThickness" Value="0,0,1,1"/>
                    <Setter Property="Padding" Value="8,6"/>
                    <Setter Property="FontWeight" Value="SemiBold"/>
                </Style>
                <Style TargetType="DataGridCell">
                    <Setter Property="BorderThickness" Value="0"/>
                    <Setter Property="Padding" Value="6,3"/>
                    <Style.Triggers>
                        <Trigger Property="IsSelected" Value="True">
                            <Setter Property="Background" Value="$($t.Selected)"/>
                            <Setter Property="Foreground" Value="$($t.SelectedText)"/>
                        </Trigger>
                    </Style.Triggers>
                </Style>
                <Style TargetType="DataGridRow">
                    <Style.Triggers>
                        <DataTrigger Binding="{Binding CanDelete}" Value="False">
                            <Setter Property="Foreground" Value="$($t.Placeholder)"/>
                        </DataTrigger>
                    </Style.Triggers>
                </Style>
            </DataGrid.Resources>
            <DataGrid.Columns>
                <DataGridTemplateColumn Header="" Width="36" SortMemberPath="IsSelected" CanUserSort="True">
                    <DataGridTemplateColumn.CellTemplate>
                        <DataTemplate>
                            <CheckBox IsChecked="{Binding IsSelected, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}"
                                      IsEnabled="{Binding CanDelete}" HorizontalAlignment="Center" Margin="0"/>
                        </DataTemplate>
                    </DataGridTemplateColumn.CellTemplate>
                </DataGridTemplateColumn>
                <DataGridTextColumn Header="Name" Binding="{Binding DisplayName}" Width="*" MinWidth="260" IsReadOnly="True"/>
                <DataGridTextColumn Header="Source" Binding="{Binding Source}" Width="85" IsReadOnly="True"/>
                <DataGridTextColumn Header="Publisher" Binding="{Binding Publisher}" Width="110" IsReadOnly="True"/>
                <DataGridTextColumn Header="Ver" Binding="{Binding Version}" Width="45" IsReadOnly="True"/>
                <DataGridTextColumn Header="Assigned" Binding="{Binding Assigned}" SortMemberPath="AssignmentCount" Width="85" IsReadOnly="True"/>
                <DataGridTextColumn Header="Created" Binding="{Binding Created, StringFormat='yyyy-MM-dd HH:mm'}" Width="115" IsReadOnly="True"/>
                <DataGridTextColumn Header="Modified" Binding="{Binding Modified, StringFormat='yyyy-MM-dd HH:mm'}" Width="115" IsReadOnly="True"/>
                <DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="105" IsReadOnly="True"/>
            </DataGrid.Columns>
        </DataGrid>

        <!-- Details for the highlighted row -->
        <Border Grid.Row="3" Margin="0,10,0,0" Padding="12,8" Height="130"
                Background="$($t.Control)" BorderBrush="$($t.Border)" BorderThickness="1"
                DataContext="{Binding SelectedItem, ElementName=Grid}">
            <Grid>
                <TextBlock Text="Click a row to see its full description." Foreground="$($t.Placeholder)" FontStyle="Italic">
                    <TextBlock.Style>
                        <Style TargetType="TextBlock">
                            <Setter Property="Visibility" Value="Collapsed"/>
                            <Style.Triggers>
                                <DataTrigger Binding="{Binding}" Value="{x:Null}">
                                    <Setter Property="Visibility" Value="Visible"/>
                                </DataTrigger>
                            </Style.Triggers>
                        </Style>
                    </TextBlock.Style>
                </TextBlock>
                <DockPanel>
                    <TextBlock DockPanel.Dock="Top" FontWeight="SemiBold" FontSize="14" Foreground="$($t.Text)" TextTrimming="CharacterEllipsis">
                        <Run Text="{Binding DisplayName, Mode=OneWay}"/>
                        <Run Text="{Binding Id, Mode=OneWay}" FontWeight="Normal" FontSize="11" Foreground="$($t.Placeholder)"/>
                    </TextBlock>
                    <TextBox Text="{Binding Description, Mode=OneWay}" IsReadOnly="True" TextWrapping="Wrap"
                             VerticalScrollBarVisibility="Auto" Margin="0,4,0,0" Padding="0"
                             Background="Transparent" BorderThickness="0" Foreground="$($t.Text2)" FontSize="13"/>
                </DockPanel>
            </Grid>
        </Border>

        <!-- Backup folder + backup buttons (also where pre-delete backups go) -->
        <DockPanel Name="BackupBar" Grid.Row="4" Margin="0,10,0,0">
            <TextBlock DockPanel.Dock="Left" Text="Back up to:" Foreground="$($t.Text2)" VerticalAlignment="Center" Margin="0,0,8,0"/>
            <Button DockPanel.Dock="Right" Name="BackupAllBtn" Content="Back up all" Margin="0" Style="{StaticResource GreenButton}"
                    ToolTip="Back up every loaded script (including Microsoft ones and any hidden by filters) to this folder. Nothing is deleted."/>
            <Button DockPanel.Dock="Right" Name="BackupCheckedBtn" Content="Back up checked" Style="{StaticResource GreenButton}"
                    ToolTip="Back up the checked scripts to this folder. Nothing is deleted, and they stay checked."/>
            <Button DockPanel.Dock="Right" Name="OpenBackupBtn" Content="Open"/>
            <Button DockPanel.Dock="Right" Name="BrowseBackupBtn" Content="Change..."/>
            <TextBox Name="BackupPathBox" Padding="6,4" Margin="0,0,8,0" VerticalContentAlignment="Center" FontSize="13"
                     ToolTip="Local folder or UNC share. Each script is backed up to its own folder here, named after the script."/>
        </DockPanel>

        <DockPanel Grid.Row="5" Margin="0,12,0,0" LastChildFill="False">
            <StackPanel DockPanel.Dock="Left" Orientation="Horizontal">
                <Button Name="CheckVisibleBtn" Content="Check all visible"/>
                <Button Name="CheckHighlightedBtn" Content="Check highlighted (Space)"/>
                <Button Name="ClearBtn" Content="Clear checks"/>
                <Button Name="RefreshBtn" Content="Refresh"/>
                <Button Name="ExportBtn" Content="Export CSV"/>
            </StackPanel>
            <StackPanel DockPanel.Dock="Right" Orientation="Horizontal">
                <TextBlock Name="StatusText" Foreground="$($t.Text2)" VerticalAlignment="Center" Margin="0,0,12,0"/>
                <!-- Red when something is checked, blue when nothing is -->
                <Button Name="DeleteBtn" Content="Delete checked" Margin="0">
                    <Button.Style>
                        <Style TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
                            <Setter Property="Foreground" Value="White"/>
                            <Setter Property="Background" Value="$($t.Red)"/>
                            <Setter Property="BorderBrush" Value="$($t.RedHover)"/>
                            <Setter Property="Template" Value="{StaticResource FlatButton}"/>
                            <Style.Triggers>
                                <Trigger Property="IsMouseOver" Value="True">
                                    <Setter Property="Background" Value="$($t.RedHover)"/>
                                </Trigger>
                                <Trigger Property="IsEnabled" Value="False">
                                    <Setter Property="Background" Value="$($t.Blue)"/>
                                    <Setter Property="BorderBrush" Value="$($t.Blue)"/>
                                </Trigger>
                            </Style.Triggers>
                        </Style>
                    </Button.Style>
                </Button>
            </StackPanel>
        </DockPanel>
    </Grid>
</Window>
"@

$window = [Windows.Markup.XamlReader]::Parse($xaml)
$ui = @{}
foreach ($n in 'SearchBox','SearchPlaceholder','HideGlobal','OnlyUnassigned','OnlySelected','CountText','Grid',
               'CheckVisibleBtn','CheckHighlightedBtn','ClearBtn','RefreshBtn','ExportBtn','StatusText','DeleteBtn',
               'BackupPathBox','BrowseBackupBtn','OpenBackupBtn','BackupCheckedBtn','BackupAllBtn') {
    $ui[$n] = $window.FindName($n)
}

# Stays visible with -NoBackup: that only skips the pre-delete backup, the backup buttons still work
$ui.BackupPathBox.Text = $script:BackupRoot

$ui.Grid.ItemsSource = $script:Rows
$script:View = [System.Windows.Data.CollectionViewSource]::GetDefaultView($script:Rows)
$script:View.SortDescriptions.Add([System.ComponentModel.SortDescription]::new('DisplayName', 'Ascending'))

$script:View.Filter = [Predicate[object]] {
    param($r)
    if ($ui.HideGlobal.IsChecked -and $r.IsGlobal) { return $false }
    if ($ui.OnlyUnassigned.IsChecked -and $r.AssignmentCount -ne 0) { return $false }
    if ($ui.OnlySelected.IsChecked -and -not $r.IsSelected) { return $false }
    $q = $ui.SearchBox.Text.Trim()
    if (-not $q) { return $true }
    # every space-separated term must match somewhere (AND search)
    $hay = "$($r.DisplayName) $($r.Description) $($r.Publisher)"
    foreach ($term in $q -split '\s+') {
        $inText = $hay.IndexOf($term, [StringComparison]::OrdinalIgnoreCase) -ge 0
        # only match against the GUID for longer terms, so "1" doesn't hit every ID
        $inId = $term.Length -ge 8 -and $r.Id.IndexOf($term, [StringComparison]::OrdinalIgnoreCase) -ge 0
        if (-not ($inText -or $inId)) { return $false }
    }
    return $true
}

function Update-Counts {
    $visible = @($script:View).Count
    $checked = @($script:Rows | Where-Object IsSelected).Count
    $ui.CountText.Text = "Showing $visible of $($script:Rows.Count)   |   $checked checked"
    $ui.DeleteBtn.Content = if ($checked) { "Delete $checked checked" } else { 'Delete checked' }
    $ui.DeleteBtn.IsEnabled = $checked -gt 0
    $ui.BackupCheckedBtn.Content = if ($checked) { "Back up $checked checked" } else { 'Back up checked' }
    $ui.BackupCheckedBtn.IsEnabled = $checked -gt 0
    $ui.BackupAllBtn.Content = "Back up all ($($script:Rows.Count))"
    $ui.BackupAllBtn.IsEnabled = $script:Rows.Count -gt 0
}

function Update-View { $script:View.Refresh(); Update-Counts }

function Pump-UI { $window.Dispatcher.Invoke([Action] {}, [System.Windows.Threading.DispatcherPriority]::Background) }

function Initialize-BackupRoot {
    # Validates the folder in the box and proves it's reachable and writable.
    # Returns the normalised path, or shows the problem and returns $null.
    try { $root = Resolve-BackupRoot $ui.BackupPathBox.Text }
    catch { [System.Windows.MessageBox]::Show($window, $_.Exception.Message, 'Backup folder', 'OK', 'Warning') | Out-Null; return $null }
    try {
        [IO.Directory]::CreateDirectory($root) | Out-Null
        $probe = Join-Path $root ".ircleanup_write_test_$PID"
        [IO.File]::WriteAllText($probe, ''); [IO.File]::Delete($probe)
    }
    catch {
        [System.Windows.MessageBox]::Show($window, "Can't write to the backup folder.`n`n$root`n`n$($_.Exception.Message)",
            'Backup folder', 'OK', 'Error') | Out-Null
        return $null
    }
    # Remember a folder typed into the box (Change... already saves its pick)
    if ($root -ne $script:BackupRoot) { $script:BackupRoot = $root; Save-BackupRoot $root }
    $ui.BackupPathBox.Text = $root
    $root
}

function Write-RunLog {
    # Saves a run's CSV log; returns a line for the summary. Never throws, so a share that drops
    # mid-run can't leave the window disabled.
    param($Log, [string]$Dir, [string]$Prefix)
    try {
        [IO.Directory]::CreateDirectory($Dir) | Out-Null
        $stamp = Get-Date -Format yyyyMMdd_HHmmss
        $path  = Join-Path $Dir "$($Prefix)_$stamp.csv"
        # Two quick runs can land in the same second; don't overwrite the first run's log
        for ($n = 2; [IO.File]::Exists($path); $n++) { $path = Join-Path $Dir "$($Prefix)_$($stamp)_$n.csv" }
        # -LiteralPath: a custom folder may contain [ ] which -Path treats as wildcards
        $Log | Export-Csv -LiteralPath $path @script:CsvArgs
        "Log: $path"
    }
    catch { "Log could not be saved: $($_.Exception.Message)" }
}

$ui.SearchBox.Add_TextChanged({
    $ui.SearchPlaceholder.Visibility = if ($ui.SearchBox.Text) { 'Collapsed' } else { 'Visible' }
    Update-View
})
foreach ($cb in $ui.HideGlobal, $ui.OnlyUnassigned, $ui.OnlySelected) {
    $cb.Add_Checked({ Update-View }); $cb.Add_Unchecked({ Update-View })
}

# Recount whenever any row checkbox is clicked
$ui.Grid.AddHandler([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent,
    [System.Windows.RoutedEventHandler] { Update-Counts })

# Space toggles the checkbox for all highlighted rows (shift/ctrl-click to highlight a range)
function Toggle-Highlighted {
    $rows = @($ui.Grid.SelectedItems | Where-Object CanDelete)
    if (-not $rows) { return }
    $target = @($rows | Where-Object { -not $_.IsSelected }).Count -gt 0
    foreach ($r in $rows) { $r.IsSelected = $target }
    Update-Counts
}
$ui.Grid.Add_PreviewKeyDown({
    param($s, $e)
    if ($e.Key -eq 'Space') { Toggle-Highlighted; $e.Handled = $true }
})
$ui.CheckHighlightedBtn.Add_Click({ Toggle-Highlighted })

$ui.CheckVisibleBtn.Add_Click({
    foreach ($r in @($script:View)) { if ($r.CanDelete) { $r.IsSelected = $true } }
    Update-Counts
})
$ui.ClearBtn.Add_Click({
    foreach ($r in $script:Rows) { $r.IsSelected = $false }
    Update-View
})

$ui.RefreshBtn.Add_Click({
    $ui.StatusText.Text = 'Refreshing...'; Pump-UI
    try {
        $fresh = @(Get-RemediationRows)
        $script:Rows.Clear()
        foreach ($r in $fresh) { $script:Rows.Add($r) }
        $ui.StatusText.Text = "Loaded $($fresh.Count)"
    }
    catch { $ui.StatusText.Text = 'Refresh failed'; [System.Windows.MessageBox]::Show("Refresh failed:`n$_", 'IRCleanup', 'OK', 'Error') | Out-Null }
    Update-View
})

$ui.ExportBtn.Add_Click({
    $dlg = [Microsoft.Win32.SaveFileDialog]::new()
    $dlg.Filter = 'CSV (*.csv)|*.csv'
    $dlg.FileName = "Remediations_$(Get-Date -Format yyyyMMdd_HHmmss).csv"
    if ($dlg.ShowDialog($window)) {
        @($script:View) | Select-Object IsSelected, DisplayName, Source, Publisher, Version, AssignmentCount, Created, Modified, Status, Description, Id |
            Export-Csv -LiteralPath $dlg.FileName @script:CsvArgs
        $ui.StatusText.Text = "Exported $(@($script:View).Count) rows"
    }
})

$ui.BrowseBackupBtn.Add_Click({
    $start = try { Resolve-BackupRoot $ui.BackupPathBox.Text } catch { $script:DefaultBackupRoot }
    $picked = Select-BackupFolder -Start $start -Owner $window
    if ($picked) {
        $ui.BackupPathBox.Text = $picked
        $script:BackupRoot = $picked
        Save-BackupRoot $picked
    }
})

$ui.OpenBackupBtn.Add_Click({
    try {
        $p = Resolve-BackupRoot $ui.BackupPathBox.Text
        if (-not (Test-Path -LiteralPath $p -PathType Container)) {
            throw "This folder doesn't exist yet. It's created on the first backup.`n`n$p"
        }
        Start-Process -FilePath explorer.exe -ArgumentList "`"$p`""
    }
    catch { [System.Windows.MessageBox]::Show($window, $_.Exception.Message, 'Backup folder', 'OK', 'Information') | Out-Null }
})

function Invoke-BulkBackup {
    # Shared by "Back up checked" and "Back up all". Backs up only; nothing is deleted or unchecked.
    param([object[]]$Targets)
    if (-not $Targets) { return }
    if (-not ($root = Initialize-BackupRoot)) { return }

    $window.IsEnabled = $false
    $log = [System.Collections.Generic.List[object]]::new()
    $ok = 0; $fail = 0; $i = 0
    foreach ($r in $targets) {
        $i++
        $ui.StatusText.Text = "Backing up $i of $($targets.Count)..."
        $r.Status = 'Backing up...'; Pump-UI
        try {
            $dir = Save-RemediationBackup -Id $r.Id -Root $root
            $r.Status = 'Backed up'; $ok++
            $log.Add([pscustomobject]@{ Time = Get-Date; Result = 'Backed up'; DisplayName = $r.DisplayName; Description = $r.Description; Id = $r.Id; Folder = $dir; Error = '' })
        }
        catch {
            $r.Status = 'Backup failed'; $fail++
            $log.Add([pscustomobject]@{ Time = Get-Date; Result = 'Failed'; DisplayName = $r.DisplayName; Description = $r.Description; Id = $r.Id; Folder = ''; Error = (Get-GraphErrorText $_) })
        }
    }
    $logLine = Write-RunLog -Log $log -Dir $root -Prefix 'BackupLog'

    $window.IsEnabled = $true
    $ui.StatusText.Text = "Backed up $ok, failed $fail"

    $summary = "Backed up: $ok`nFailed: $fail`n`nFolder: $root`n$logLine"
    if ($fail) {
        $summary += "`n`nFailures:`n" + (($log | Where-Object Result -eq 'Failed' | Select-Object -First 10 |
            ForEach-Object { "  - $($_.DisplayName): $($_.Error)" }) -join "`n")
    }
    [System.Windows.MessageBox]::Show($window, $summary, 'Backup complete', 'OK',
        $(if ($fail) { 'Warning' } else { 'Information' })) | Out-Null
}

$ui.BackupCheckedBtn.Add_Click({ Invoke-BulkBackup @($script:Rows | Where-Object IsSelected) })
$ui.BackupAllBtn.Add_Click({ Invoke-BulkBackup @($script:Rows) })

$ui.DeleteBtn.Add_Click({
    $targets = @($script:Rows | Where-Object { $_.IsSelected -and $_.CanDelete })
    if (-not $targets) { return }

    # Check the backup folder before asking, so a bad path never gets as far as "Yes"
    if (-not $NoBackup -and -not ($root = Initialize-BackupRoot)) { return }

    $assigned = @($targets | Where-Object { $_.AssignmentCount -gt 0 }).Count
    $preview  = ($targets | Select-Object -First 15 | ForEach-Object { "  - $($_.DisplayName)" }) -join "`n"
    if ($targets.Count -gt 15) { $preview += "`n  ... and $($targets.Count - 15) more" }
    $msg = "Permanently delete $($targets.Count) remediation script(s)?`n`n$preview"
    if ($assigned) { $msg += "`n`nWARNING: $assigned of these are still assigned to groups." }
    $msg += if ($NoBackup) { "`n`nNo backup will be taken (-NoBackup)." } else { "`n`nEach script is backed up first to:`n$root" }

    if ([System.Windows.MessageBox]::Show($window, $msg, 'Confirm bulk delete', 'YesNo', 'Warning') -ne 'Yes') { return }

    $window.IsEnabled = $false
    $log = [System.Collections.Generic.List[object]]::new()
    $ok = 0; $fail = 0; $i = 0
    foreach ($r in $targets) {
        $i++
        $ui.StatusText.Text = "Deleting $i of $($targets.Count)..."
        $r.Status = 'Deleting...'; Pump-UI
        $result = 'Deleted'; $err = ''
        try {
            if (-not $NoBackup) {
                try   { $null = Save-RemediationBackup -Id $r.Id -Root $root }
                catch { throw "Backup failed, not deleted: $(Get-GraphErrorText $_)" }
            }
            Invoke-Graph -Method DELETE -Uri "$script:GraphBaseUrl/deviceManagement/deviceHealthScripts/$($r.Id)" | Out-Null
            $r.Status = 'Deleted'; $r.IsSelected = $false; $ok++
        }
        catch {
            $result = 'Failed'; $err = Get-GraphErrorText $_
            $r.Status = 'Failed'; $fail++
        }
        $log.Add([pscustomobject]@{ Time = Get-Date; Result = $result; DisplayName = $r.DisplayName; Description = $r.Description; Id = $r.Id; Error = $err })
    }

    # Remove successfully deleted rows from the grid
    foreach ($r in @($script:Rows | Where-Object Status -eq 'Deleted')) { [void]$script:Rows.Remove($r) }

    $logDir  = if ($NoBackup) { Join-Path $env:APPDATA 'IRCleanup' } else { $root }
    $logLine = Write-RunLog -Log $log -Dir $logDir -Prefix 'DeleteLog'

    $window.IsEnabled = $true
    Update-View
    $ui.StatusText.Text = "Deleted $ok, failed $fail"

    $summary = "Deleted: $ok`nFailed: $fail`n`n$logLine"
    if (-not $NoBackup) { $summary += "`nBackups: $root" }
    if ($fail) {
        $summary += "`n`nFailures:`n" + (($log | Where-Object Result -eq 'Failed' | Select-Object -First 10 |
            ForEach-Object { "  - $($_.DisplayName): $($_.Error)" }) -join "`n")
    }
    [System.Windows.MessageBox]::Show($window, $summary, 'Bulk delete complete', 'OK',
        $(if ($fail) { 'Warning' } else { 'Information' })) | Out-Null
})

Update-View
$window.Add_ContentRendered({ $ui.SearchBox.Focus() | Out-Null })
[void]$window.ShowDialog()
