<#
.SYNOPSIS
    IRCleanup - Intune Remediation Cleanup.
    Search, sort and bulk-delete Intune Remediation scripts (deviceHealthScripts) from a WPF grid.

.DESCRIPTION
    Standalone companion to IROD. Uses the same Graph pattern (Invoke-MgGraphRequest against
    beta/deviceManagement/deviceHealthScripts with @odata.nextLink paging) and the same
    IROD_CLIENTID / IROD_TENANTID environment variables if present.

    Before each delete, the full script (detection + remediation content, settings, assignments)
    is backed up as JSON + .ps1 files under %APPDATA%\IRCleanup\Backups\<timestamp>\.

.PARAMETER ClientId
    App registration client ID. Falls back to $env:IROD_CLIENTID.

.PARAMETER TenantId
    Tenant ID. Falls back to $env:IROD_TENANTID.

.PARAMETER NoBackup
    Skip the pre-delete backup. Not recommended.

.PARAMETER Theme
    Dark or Light. Defaults to the IROD theme setting, else Dark.

.NOTES
    Required delegated permission: DeviceManagementConfiguration.ReadWrite.All
    (IROD's app registration only has .Read.All - add ReadWrite if you reuse it.)
#>
[CmdletBinding()]
param(
    [string]$ClientId,
    [string]$TenantId,
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
$script:BackupRoot   = Join-Path $env:APPDATA 'IRCleanup\Backups'

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

# --- Theme (matches IROD palette) -------------------------------------------
if (-not $Theme) {
    $Theme = 'Dark'
    $irodSettings = Join-Path $env:APPDATA 'IROD\settings.json'
    if (Test-Path $irodSettings) {
        try { if ((Get-Content $irodSettings -Raw | ConvertFrom-Json).Theme -eq 'Light') { $Theme = 'Light' } } catch { }
    }
}
$t = if ($Theme -eq 'Light') {
    @{ Window = '#F5F5F5'; Control = '#FFFFFF'; Alt = '#F7F7F7'; Hover = '#E5E5E5'; Text = '#1E1E1E'; Text2 = '#666666'
       Placeholder = '#999999'; Border = '#CCCCCC'; Header = '#E8E8E8'; Selected = '#0078D4'; SelectedText = '#FFFFFF'
       Button = '#E1E1E1'; ButtonHover = '#D5D5D5'; Green = '#2E8B57'; GreenHover = '#3E9B67'; Red = '#C42B1C'; RedHover = '#D43B2C' }
} else {
    @{ Window = '#1E1E1E'; Control = '#2D2D30'; Alt = '#2A2A2D'; Hover = '#3E3E42'; Text = '#FFFFFF'; Text2 = '#CCCCCC'
       Placeholder = '#808080'; Border = '#3F3F46'; Header = '#2D2D30'; Selected = '#094771'; SelectedText = '#FFFFFF'
       Button = '#2D2D30'; ButtonHover = '#3E3E42'; Green = '#2E8B57'; GreenHover = '#3E9B67'; Red = '#B3261E'; RedHover = '#D43B2C' }
}

# --- Graph helpers (same pattern as IROD) -----------------------------------
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
    $p = @{ Scopes = @('DeviceManagementConfiguration.ReadWrite.All'); NoWelcome = $true }
    if ($ClientId) { $p.ClientId = $ClientId }
    if ($TenantId) { $p.TenantId = $TenantId }
    try {
        Connect-MgGraph @p -ErrorAction Stop -WarningAction SilentlyContinue | Out-Null
        $ctx = Get-MgContext
        if ($ctx.Scopes -notcontains 'DeviceManagementConfiguration.ReadWrite.All') {
            Write-Host "Warning: token lacks DeviceManagementConfiguration.ReadWrite.All - deletes will fail." -ForegroundColor Yellow
        }
        Write-Host "Connected as: $($ctx.Account)" -ForegroundColor Green
        return $true
    }
    catch { Write-Host "Auth failed: $_" -ForegroundColor Red; return $false }
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

function Save-RemediationBackup {
    param([string]$Id, [string]$Folder)
    $full = Invoke-Graph -Uri "$script:GraphBaseUrl/deviceManagement/deviceHealthScripts/$Id`?`$expand=assignments"
    $safe = ($full.displayName -replace '[\\/:*?"<>|]', '_').Trim()
    $dir  = Join-Path $Folder "$safe [$Id]"
    New-Item -ItemType Directory -Path $dir -Force | Out-Null
    # .NET write: folder name contains [id], which PowerShell cmdlets would treat as a wildcard
    [IO.File]::WriteAllText((Join-Path $dir 'script.json'), ($full | ConvertTo-Json -Depth 20), [Text.UTF8Encoding]::new($false))
    foreach ($pair in @(@('detectionScriptContent', 'detect.ps1'), @('remediationScriptContent', 'remediate.ps1'))) {
        if ($full[$pair[0]]) {
            $bytes = [Convert]::FromBase64String($full[$pair[0]])
            [IO.File]::WriteAllBytes((Join-Path $dir $pair[1]), $bytes)
        }
    }
}

# --- Start -------------------------------------------------------------------
if ([string]::IsNullOrWhiteSpace($ClientId)) { $ClientId = $env:IROD_CLIENTID }
if ([string]::IsNullOrWhiteSpace($TenantId)) { $TenantId = $env:IROD_TENANTID }

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
        <Style TargetType="Button">
            <Setter Property="Background" Value="$($t.Button)"/>
            <Setter Property="Foreground" Value="$($t.Text)"/>
            <Setter Property="BorderBrush" Value="$($t.Border)"/>
            <Setter Property="Padding" Value="12,6"/>
            <Setter Property="Margin" Value="0,0,8,0"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Cursor" Value="Hand"/>
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
                <DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="90" IsReadOnly="True"/>
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

        <DockPanel Grid.Row="4" Margin="0,12,0,0" LastChildFill="False">
            <StackPanel DockPanel.Dock="Left" Orientation="Horizontal">
                <Button Name="CheckVisibleBtn" Content="Check all visible"/>
                <Button Name="CheckHighlightedBtn" Content="Check highlighted (Space)"/>
                <Button Name="ClearBtn" Content="Clear checks"/>
                <Button Name="RefreshBtn" Content="Refresh"/>
                <Button Name="ExportBtn" Content="Export CSV"/>
            </StackPanel>
            <StackPanel DockPanel.Dock="Right" Orientation="Horizontal">
                <TextBlock Name="StatusText" Foreground="$($t.Text2)" VerticalAlignment="Center" Margin="0,0,12,0"/>
                <Button Name="DeleteBtn" Content="Delete checked" Margin="0" Foreground="White" BorderBrush="$($t.RedHover)" Background="$($t.Red)"/>
            </StackPanel>
        </DockPanel>
    </Grid>
</Window>
"@

$window = [Windows.Markup.XamlReader]::Parse($xaml)
$ui = @{}
foreach ($n in 'SearchBox','SearchPlaceholder','HideGlobal','OnlyUnassigned','OnlySelected','CountText','Grid',
               'CheckVisibleBtn','CheckHighlightedBtn','ClearBtn','RefreshBtn','ExportBtn','StatusText','DeleteBtn') {
    $ui[$n] = $window.FindName($n)
}

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
}

function Update-View { $script:View.Refresh(); Update-Counts }

function Pump-UI { $window.Dispatcher.Invoke([Action] {}, [System.Windows.Threading.DispatcherPriority]::Background) }

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
            Export-Csv -Path $dlg.FileName -NoTypeInformation -Encoding UTF8
        $ui.StatusText.Text = "Exported $(@($script:View).Count) rows"
    }
})

$ui.DeleteBtn.Add_Click({
    $targets = @($script:Rows | Where-Object { $_.IsSelected -and $_.CanDelete })
    if (-not $targets) { return }

    $assigned = @($targets | Where-Object { $_.AssignmentCount -gt 0 }).Count
    $preview  = ($targets | Select-Object -First 15 | ForEach-Object { "  - $($_.DisplayName)" }) -join "`n"
    if ($targets.Count -gt 15) { $preview += "`n  ... and $($targets.Count - 15) more" }
    $msg = "Permanently delete $($targets.Count) remediation script(s)?`n`n$preview"
    if ($assigned) { $msg += "`n`nWARNING: $assigned of these are still assigned to groups." }
    $msg += if ($NoBackup) { "`n`nNo backup will be taken (-NoBackup)." } else { "`n`nA backup of each script is saved first." }

    if ([System.Windows.MessageBox]::Show($window, $msg, 'Confirm bulk delete', 'YesNo', 'Warning') -ne 'Yes') { return }

    $backupDir = Join-Path $script:BackupRoot (Get-Date -Format 'yyyyMMdd_HHmmss')
    if (-not $NoBackup) { New-Item -ItemType Directory -Path $backupDir -Force | Out-Null }

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
                try   { Save-RemediationBackup -Id $r.Id -Folder $backupDir }
                catch { throw "Backup failed, not deleted: $(Get-GraphErrorText $_)" }
            }
            Invoke-Graph -Method DELETE -Uri "$script:GraphBaseUrl/deviceManagement/deviceHealthScripts/$($r.Id)" | Out-Null
            $r.Status = 'Deleted'; $r.IsSelected = $false; $ok++
        }
        catch {
            $result = 'Failed'; $err = Get-GraphErrorText $_
            $r.Status = 'Failed'; $fail++
        }
        $log.Add([pscustomobject]@{ Time = Get-Date; Result = $result; DisplayName = $r.DisplayName; Id = $r.Id; Error = $err })
    }

    # Remove successfully deleted rows from the grid
    foreach ($r in @($script:Rows | Where-Object Status -eq 'Deleted')) { [void]$script:Rows.Remove($r) }

    $logDir = if ($NoBackup) { Join-Path $env:APPDATA 'IRCleanup' } else { $backupDir }
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    $logPath = Join-Path $logDir "DeleteLog_$(Get-Date -Format yyyyMMdd_HHmmss).csv"
    $log | Export-Csv -Path $logPath -NoTypeInformation -Encoding UTF8

    $window.IsEnabled = $true
    Update-View
    $ui.StatusText.Text = "Deleted $ok, failed $fail"

    $summary = "Deleted: $ok`nFailed: $fail`n`nLog: $logPath"
    if (-not $NoBackup) { $summary += "`nBackups: $backupDir" }
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
