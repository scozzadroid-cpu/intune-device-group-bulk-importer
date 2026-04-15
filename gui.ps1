#Requires -Version 5.1
# ============================================================
#  Intune Device Group Bulk Importer — GUI
#  WPF + Runspace async — UI never freezes during Graph calls
# ============================================================

# Bypass execution policy for this process
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force -ErrorAction SilentlyContinue

# Disable WAM before any MSAL module loads (prevents window-handle error)
[System.Environment]::SetEnvironmentVariable("MSAL_DISABLE_WAM", "1", "Process")
$env:MSAL_DISABLE_WAM = "1"

# STA check (WPF requires STA thread)
if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -STA -NoProfile -File `"$PSCommandPath`"" -Wait
    exit
}

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms

# ── XAML ─────────────────────────────────────────────────────
[xml]$Xaml = @'
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Intune Bulk Group Importer"
    Height="760" Width="900"
    MinHeight="620" MinWidth="720"
    WindowStartupLocation="CenterScreen"
    Background="#0F0F1C"
    FontFamily="Segoe UI">

  <Window.Resources>

    <Style x:Key="BtnPrimary" TargetType="Button">
      <Setter Property="Foreground" Value="White"/>
      <Setter Property="Padding" Value="16,8"/>
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="bd" Background="{TemplateBinding Background}"
                    CornerRadius="6" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsEnabled" Value="False">
                <Setter TargetName="bd" Property="Opacity" Value="0.30"/>
              </Trigger>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="bd" Property="Opacity" Value="0.80"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="Btn" TargetType="Button" BasedOn="{StaticResource BtnPrimary}">
      <Setter Property="FontWeight" Value="Normal"/>
    </Style>

    <Style x:Key="Txt" TargetType="TextBox">
      <Setter Property="Background" Value="#12121F"/>
      <Setter Property="Foreground" Value="#E8E8F8"/>
      <Setter Property="BorderBrush" Value="#2E2E50"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="CaretBrush" Value="#7C5CFC"/>
      <Setter Property="SelectionBrush" Value="#4A3A90"/>
      <Setter Property="FontFamily" Value="Consolas"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="Padding" Value="9,7"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="TextBox">
            <Border x:Name="bd"
                    Background="{TemplateBinding Background}"
                    BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}"
                    CornerRadius="6">
              <ScrollViewer x:Name="PART_ContentHost" Margin="{TemplateBinding Padding}"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsFocused" Value="True">
                <Setter TargetName="bd" Property="BorderBrush" Value="#7C5CFC"/>
                <Setter TargetName="bd" Property="Effect">
                  <Setter.Value>
                    <DropShadowEffect Color="#7C5CFC" BlurRadius="8" Opacity="0.25" ShadowDepth="0"/>
                  </Setter.Value>
                </Setter>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

  </Window.Resources>

  <Grid>

    <!-- ── Main UI ── -->
    <Grid x:Name="mainGrid" Margin="22" IsEnabled="False">
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="*"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="210"/>
      </Grid.RowDefinitions>

      <!-- Header -->
      <StackPanel Grid.Row="0" Margin="0,0,0,16">
        <TextBlock Text="Intune Bulk Group Importer"
                   FontSize="21" FontWeight="Bold" Foreground="#EEEEFF"/>
        <TextBlock Text="Resolve devices via Microsoft Graph and generate the Intune group import CSV"
                   FontSize="11" Foreground="#6060A0" Margin="0,4,0,0"/>
      </StackPanel>

      <!-- Auth bar -->
      <Border Grid.Row="1"
              Background="#181830"
              BorderBrush="#2A2A50" BorderThickness="1"
              CornerRadius="8" Padding="16,12" Margin="0,0,0,16">
        <Grid>
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="Auto"/>
          </Grid.ColumnDefinitions>
          <StackPanel Grid.Column="0" VerticalAlignment="Center">
            <TextBlock x:Name="txtAuthStatus" Text="Not connected"
                       FontSize="13" FontWeight="SemiBold" Foreground="#6060A0"/>
            <TextBlock x:Name="txtAuthDetail" Text=""
                       FontSize="10" Foreground="#404080" Margin="0,3,0,0"/>
            <StackPanel Orientation="Horizontal" Margin="0,7,0,0">
              <RadioButton x:Name="rbDeviceCode" Content="Device code" IsChecked="True"
                           Foreground="#9090C0" FontSize="11" Margin="0,0,16,0"
                           GroupName="AuthMethod"/>
              <RadioButton x:Name="rbInteractive" Content="Interactive (browser popup)"
                           Foreground="#9090C0" FontSize="11"
                           GroupName="AuthMethod"/>
            </StackPanel>
          </StackPanel>
          <Button x:Name="btnDisconnect" Grid.Column="1"
                  Content="Disconnect" Background="#3A1F5E"
                  Style="{StaticResource Btn}" Margin="0,0,10,0"
                  Visibility="Collapsed"/>
          <Button x:Name="btnConnect" Grid.Column="2"
                  Content="Connect to Microsoft Graph"
                  Background="#7C5CFC" Style="{StaticResource BtnPrimary}"/>
        </Grid>
      </Border>

      <!-- Tabs -->
      <TabControl x:Name="tabMain" Grid.Row="2"
                  Background="#141428"
                  BorderBrush="#2A2A50" BorderThickness="1">
        <TabControl.Resources>
          <Style TargetType="TabItem">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="#5858A0"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Padding" Value="20,10"/>
            <Setter Property="Template">
              <Setter.Value>
                <ControlTemplate TargetType="TabItem">
                  <Border x:Name="bd" Background="Transparent"
                          BorderThickness="0,0,0,2" BorderBrush="Transparent"
                          Padding="{TemplateBinding Padding}">
                    <TextBlock Text="{TemplateBinding Header}"
                               Foreground="{TemplateBinding Foreground}"/>
                  </Border>
                  <ControlTemplate.Triggers>
                    <Trigger Property="IsSelected" Value="True">
                      <Setter TargetName="bd" Property="BorderBrush" Value="#7C5CFC"/>
                      <Setter Property="Foreground" Value="#EEEEFF"/>
                    </Trigger>
                    <Trigger Property="IsMouseOver" Value="True">
                      <Setter Property="Foreground" Value="#A0A0D8"/>
                    </Trigger>
                  </ControlTemplate.Triggers>
                </ControlTemplate>
              </Setter.Value>
            </Setter>
          </Style>
        </TabControl.Resources>

        <!-- Tab: Hostname -->
        <TabItem Header="By Hostname">
          <Grid Margin="18">
            <Grid.RowDefinitions>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="*"/>
              <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>
            <Grid Grid.Row="0" Margin="0,0,0,10">
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>
              <TextBox x:Name="txtHostnameFile" Grid.Column="0"
                       Style="{StaticResource Txt}" Text="C:\TEMP\pc.txt"/>
              <Button x:Name="btnBrowseHostname" Grid.Column="1"
                      Content="Browse…" Background="#22223A"
                      Style="{StaticResource Btn}" Margin="8,0,0,0"/>
              <Button x:Name="btnLoadHostname" Grid.Column="2"
                      Content="Load" Background="#22223A"
                      Style="{StaticResource Btn}" Margin="6,0,0,0"/>
            </Grid>
            <Grid Grid.Row="1">
              <TextBox x:Name="txtHostnameList"
                       Style="{StaticResource Txt}"
                       AcceptsReturn="True" TextWrapping="NoWrap"
                       VerticalScrollBarVisibility="Auto"
                       HorizontalScrollBarVisibility="Auto"
                       Background="#0C0C18"/>
              <TextBlock x:Name="phHostname"
                         Text="Paste hostnames here, one per line&#x0a;&#x0a;Example:&#x0a;LAPTOP-001&#x0a;DESKTOP-FINANCE-03&#x0a;WS-HR-12"
                         Foreground="#252545" FontFamily="Consolas" FontSize="12"
                         Margin="12,10,0,0" IsHitTestVisible="False" VerticalAlignment="Top"/>
            </Grid>
            <Grid Grid.Row="2" Margin="0,12,0,0">
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>
              <TextBlock Text="Output:" Foreground="#6060A0" FontSize="12"
                         VerticalAlignment="Center" Margin="0,0,10,0"/>
              <TextBox x:Name="txtHostnameOutput" Grid.Column="1"
                       Style="{StaticResource Txt}"
                       Text="C:\TEMP\IntuneGroupImport.csv"/>
              <Button x:Name="btnBrowseHostnameOut" Grid.Column="2"
                      Content="…" Background="#22223A"
                      Style="{StaticResource Btn}" Width="36" Margin="8,0"/>
              <Button x:Name="btnRunHostname" Grid.Column="3"
                      Content="  Run Import  " Background="#7C5CFC"
                      Style="{StaticResource BtnPrimary}" IsEnabled="False"/>
            </Grid>
          </Grid>
        </TabItem>

        <!-- Tab: Serial -->
        <TabItem Header="By Serial Number">
          <Grid Margin="18">
            <Grid.RowDefinitions>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="*"/>
              <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>
            <Grid Grid.Row="0" Margin="0,0,0,10">
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>
              <TextBox x:Name="txtSerialFile" Grid.Column="0"
                       Style="{StaticResource Txt}" Text="C:\TEMP\serials.txt"/>
              <Button x:Name="btnBrowseSerial" Grid.Column="1"
                      Content="Browse…" Background="#22223A"
                      Style="{StaticResource Btn}" Margin="8,0,0,0"/>
              <Button x:Name="btnLoadSerial" Grid.Column="2"
                      Content="Load" Background="#22223A"
                      Style="{StaticResource Btn}" Margin="6,0,0,0"/>
            </Grid>
            <Grid Grid.Row="1">
              <TextBox x:Name="txtSerialList"
                       Style="{StaticResource Txt}"
                       AcceptsReturn="True" TextWrapping="NoWrap"
                       VerticalScrollBarVisibility="Auto"
                       HorizontalScrollBarVisibility="Auto"
                       Background="#0C0C18"/>
              <TextBlock x:Name="phSerial"
                         Text="Paste serial numbers here, one per line&#x0a;&#x0a;Example:&#x0a;C02XG2JHJTD5&#x0a;5CG1234ABC&#x0a;VMware-56 4d 3f 21"
                         Foreground="#252545" FontFamily="Consolas" FontSize="12"
                         Margin="12,10,0,0" IsHitTestVisible="False" VerticalAlignment="Top"/>
            </Grid>
            <Grid Grid.Row="2" Margin="0,12,0,0">
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>
              <TextBlock Text="Output:" Foreground="#6060A0" FontSize="12"
                         VerticalAlignment="Center" Margin="0,0,10,0"/>
              <TextBox x:Name="txtSerialOutput" Grid.Column="1"
                       Style="{StaticResource Txt}"
                       Text="C:\TEMP\IntuneGroupImport.csv"/>
              <Button x:Name="btnBrowseSerialOut" Grid.Column="2"
                      Content="…" Background="#22223A"
                      Style="{StaticResource Btn}" Width="36" Margin="8,0"/>
              <Button x:Name="btnRunSerial" Grid.Column="3"
                      Content="  Run Import  " Background="#7C5CFC"
                      Style="{StaticResource BtnPrimary}" IsEnabled="False"/>
            </Grid>
          </Grid>
        </TabItem>

      </TabControl>

      <!-- Status + Progress -->
      <StackPanel Grid.Row="3" Margin="0,12,0,8">
        <Grid>
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
          </Grid.ColumnDefinitions>
          <TextBlock x:Name="txtStatus" Text="Ready"
                     Foreground="#5858A0" FontSize="11" VerticalAlignment="Center"/>
          <TextBlock x:Name="txtProgressLabel" Grid.Column="1" Text=""
                     Foreground="#7C5CFC" FontSize="11" FontWeight="SemiBold"
                     VerticalAlignment="Center"/>
        </Grid>
        <ProgressBar x:Name="progressBar" Height="3" Margin="0,6,0,0"
                     Background="#181830" Foreground="#7C5CFC"
                     BorderThickness="0" Value="0" Maximum="100"
                     IsIndeterminate="False"/>
      </StackPanel>

      <!-- Log panel -->
      <Border Grid.Row="4" Background="#090914"
              BorderBrush="#1E1E38" BorderThickness="1"
              CornerRadius="8">
        <Grid>
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
          </Grid.RowDefinitions>
          <Grid Margin="12,7,12,5">
            <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
              <TextBlock Text="LOG" Foreground="#353565" FontSize="10"
                         FontWeight="Bold" VerticalAlignment="Center"/>
            </StackPanel>
            <Button x:Name="btnClearLog" Content="Clear"
                    HorizontalAlignment="Right" VerticalAlignment="Center"
                    Background="#1A1A32" Style="{StaticResource Btn}"
                    Padding="10,3" FontSize="10"/>
          </Grid>
          <TextBox x:Name="txtLog" Grid.Row="1"
                   Background="Transparent"
                   Foreground="#8080B8"
                   BorderThickness="0"
                   FontFamily="Consolas" FontSize="11"
                   IsReadOnly="True" TextWrapping="NoWrap"
                   VerticalScrollBarVisibility="Auto"
                   HorizontalScrollBarVisibility="Auto"
                   Padding="12,2,12,10"/>
        </Grid>
      </Border>

    </Grid>

    <!-- ── Prerequisites overlay ── -->
    <Border x:Name="setupOverlay" Background="#CC0D0D1A"
            Panel.ZIndex="100" Visibility="Collapsed">
      <Border Background="#181830"
              BorderBrush="#2A2A50" BorderThickness="1"
              CornerRadius="12" Padding="32"
              Width="520" MaxHeight="580"
              HorizontalAlignment="Center" VerticalAlignment="Center">
        <StackPanel>

          <StackPanel Margin="0,0,0,8">
            <TextBlock Text="Prerequisites Required"
                       FontSize="18" FontWeight="Bold" Foreground="#EEEEFF"/>
            <TextBlock Text="One or more required PowerShell modules are not installed. Click Install to set them up automatically."
                       FontSize="12" Foreground="#6868A8" TextWrapping="Wrap" Margin="0,8,0,0"/>
          </StackPanel>

          <Border Background="#0F0F20"
                  BorderBrush="#252545" BorderThickness="1"
                  CornerRadius="8" Padding="18,14" Margin="0,16,0,0">
            <StackPanel>

              <Grid Margin="0,0,0,12">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="14"/>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <Ellipse x:Name="dot0" Width="9" Height="9" Fill="#3A3A70" VerticalAlignment="Center"/>
                <TextBlock Grid.Column="1" Text="Microsoft.Graph.Authentication"
                           Foreground="#B8B8E0" FontSize="12" FontFamily="Consolas" Margin="12,0,0,0"/>
                <TextBlock x:Name="modStatus0" Grid.Column="2" Text="—"
                           Foreground="#4A4A80" FontSize="11" VerticalAlignment="Center"/>
              </Grid>

              <Grid Margin="0,0,0,12">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="14"/>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <Ellipse x:Name="dot1" Width="9" Height="9" Fill="#3A3A70" VerticalAlignment="Center"/>
                <TextBlock Grid.Column="1" Text="Microsoft.Graph.DeviceManagement"
                           Foreground="#B8B8E0" FontSize="12" FontFamily="Consolas" Margin="12,0,0,0"/>
                <TextBlock x:Name="modStatus1" Grid.Column="2" Text="—"
                           Foreground="#4A4A80" FontSize="11" VerticalAlignment="Center"/>
              </Grid>

              <Grid>
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="14"/>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <Ellipse x:Name="dot2" Width="9" Height="9" Fill="#3A3A70" VerticalAlignment="Center"/>
                <TextBlock Grid.Column="1" Text="Microsoft.Graph.Identity.DirectoryManagement"
                           Foreground="#B8B8E0" FontSize="12" FontFamily="Consolas" Margin="12,0,0,0"/>
                <TextBlock x:Name="modStatus2" Grid.Column="2" Text="—"
                           Foreground="#4A4A80" FontSize="11" VerticalAlignment="Center"/>
              </Grid>

            </StackPanel>
          </Border>

          <ProgressBar x:Name="prereqProgress" Height="3" Margin="0,16,0,0"
                       Background="#0F0F20" Foreground="#7C5CFC" BorderThickness="0"
                       IsIndeterminate="True" Visibility="Collapsed"/>

          <TextBox x:Name="prereqLog" Height="110" Margin="0,10,0,0"
                   Background="#0C0C1A" Foreground="#8888B8"
                   BorderBrush="#252545" BorderThickness="1"
                   FontFamily="Consolas" FontSize="11" IsReadOnly="True"
                   VerticalScrollBarVisibility="Auto"
                   TextWrapping="NoWrap" Padding="10,8"
                   Visibility="Collapsed"/>

          <Grid Margin="0,20,0,0">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="Auto"/>
              <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <TextBlock x:Name="prereqNote" Grid.Column="0"
                       Text="Modules will be installed for the current user only."
                       Foreground="#454570" FontSize="10"
                       VerticalAlignment="Center" TextWrapping="Wrap"/>
            <Button x:Name="btnInstall" Grid.Column="1"
                    Content="Install" Background="#7C5CFC"
                    Style="{StaticResource BtnPrimary}" Margin="0,0,10,0"/>
            <Button x:Name="btnContinue" Grid.Column="2"
                    Content="Continue  →" Background="#00B896"
                    Style="{StaticResource BtnPrimary}" Visibility="Collapsed"/>
          </Grid>

        </StackPanel>
      </Border>
    </Border>

  </Grid>
</Window>
'@

$Xaml.Window.RemoveAttribute("x:Class")
$Reader = [System.Xml.XmlNodeReader]::new($Xaml)
try {
    $window = [Windows.Markup.XamlReader]::Load($Reader)
} catch {
    [System.Windows.MessageBox]::Show("XAML load failed:`n$_","Fatal Error","OK","Error") | Out-Null
    exit 1
}

# ── Find controls ─────────────────────────────────────────────
$mainGrid            = $window.FindName("mainGrid")
$setupOverlay        = $window.FindName("setupOverlay")

$dot0                = $window.FindName("dot0")
$dot1                = $window.FindName("dot1")
$dot2                = $window.FindName("dot2")
$modStatus0          = $window.FindName("modStatus0")
$modStatus1          = $window.FindName("modStatus1")
$modStatus2          = $window.FindName("modStatus2")
$prereqProgress      = $window.FindName("prereqProgress")
$prereqLog           = $window.FindName("prereqLog")
$prereqNote          = $window.FindName("prereqNote")
$btnInstall          = $window.FindName("btnInstall")
$btnContinue         = $window.FindName("btnContinue")

$txtAuthStatus       = $window.FindName("txtAuthStatus")
$txtAuthDetail       = $window.FindName("txtAuthDetail")
$btnConnect          = $window.FindName("btnConnect")
$btnDisconnect       = $window.FindName("btnDisconnect")
$rbDeviceCode        = $window.FindName("rbDeviceCode")
$rbInteractive       = $window.FindName("rbInteractive")

$txtHostnameFile     = $window.FindName("txtHostnameFile")
$btnBrowseHostname   = $window.FindName("btnBrowseHostname")
$btnLoadHostname     = $window.FindName("btnLoadHostname")
$txtHostnameList     = $window.FindName("txtHostnameList")
$txtHostnameOutput   = $window.FindName("txtHostnameOutput")
$btnBrowseHostnameOut= $window.FindName("btnBrowseHostnameOut")
$btnRunHostname      = $window.FindName("btnRunHostname")

$phHostname          = $window.FindName("phHostname")
$phSerial            = $window.FindName("phSerial")

$txtSerialFile       = $window.FindName("txtSerialFile")
$btnBrowseSerial     = $window.FindName("btnBrowseSerial")
$btnLoadSerial       = $window.FindName("btnLoadSerial")
$txtSerialList       = $window.FindName("txtSerialList")
$txtSerialOutput     = $window.FindName("txtSerialOutput")
$btnBrowseSerialOut  = $window.FindName("btnBrowseSerialOut")
$btnRunSerial        = $window.FindName("btnRunSerial")

$txtStatus           = $window.FindName("txtStatus")
$txtProgressLabel    = $window.FindName("txtProgressLabel")
$progressBar         = $window.FindName("progressBar")
$txtLog              = $window.FindName("txtLog")
$btnClearLog         = $window.FindName("btnClearLog")

# ── Thread-safe shared state ──────────────────────────────────
$ui = [hashtable]::Synchronized(@{
    Window        = $window
    Log           = $txtLog
    Status        = $txtStatus
    Progress      = $progressBar
    ProgressLabel = $txtProgressLabel
    BtnRunH       = $btnRunHostname
    BtnRunS       = $btnRunSerial
    IsRunning     = $false
})

# ── Prerequisites helpers ─────────────────────────────────────
$script:RequiredModules = @(
    "Microsoft.Graph.Authentication",
    "Microsoft.Graph.DeviceManagement",
    "Microsoft.Graph.Identity.DirectoryManagement"
)

$script:DotControls    = @($dot0, $dot1, $dot2)
$script:StatusControls = @($modStatus0, $modStatus1, $modStatus2)

function Test-Prerequisites {
    $allOk = $true
    for ($i = 0; $i -lt $script:RequiredModules.Count; $i++) {
        $ok    = [bool](Get-Module -ListAvailable -Name $script:RequiredModules[$i])
        $color = if ($ok) { '#00C896' } else { '#FF4757' }
        $script:DotControls[$i].Fill          = $color
        $script:StatusControls[$i].Text       = if ($ok) { 'Installed' } else { 'Missing' }
        $script:StatusControls[$i].Foreground = $color
        if (!$ok) { $allOk = $false }
    }
    return $allOk
}

# ── Startup: check prerequisites ──────────────────────────────
$window.Add_Loaded({
    $allOk = Test-Prerequisites
    if ($allOk) {
        $mainGrid.IsEnabled = $true
        Import-Module Microsoft.Graph.Authentication -ErrorAction SilentlyContinue
    } else {
        $setupOverlay.Visibility = 'Visible'
    }
})

# ── Install prerequisites worker ──────────────────────────────
$prereqWorker = {
    param($uiSetup, $modules)

    function PLog([string]$m) {
        $t = Get-Date -Format "HH:mm:ss"
        $uiSetup.Window.Dispatcher.Invoke([Action]{
            $uiSetup.Log.AppendText("[$t] $m`n")
            $uiSetup.Log.ScrollToEnd()
        })
    }

    try {
        PLog "Checking NuGet provider…"
        $nuget = Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue
        if (!$nuget -or $nuget.Version -lt [Version]"2.8.5.201") {
            PLog "Installing NuGet provider…"
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser | Out-Null
            PLog "NuGet provider installed."
        } else {
            PLog "NuGet provider OK."
        }

        foreach ($mod in $modules) {
            if (Get-Module -ListAvailable -Name $mod) {
                PLog "$mod — already installed."
            } else {
                PLog "Installing $mod…"
                Install-Module -Name $mod -Force -AllowClobber -Scope CurrentUser -Repository PSGallery
                PLog "$mod installed."
            }
        }

        PLog "All prerequisites ready."

        $uiSetup.Window.Dispatcher.Invoke([Action]{
            for ($i = 0; $i -lt $uiSetup.DotControls.Count; $i++) {
                $uiSetup.DotControls[$i].Fill          = '#00C896'
                $uiSetup.StatusControls[$i].Text       = 'Installed'
                $uiSetup.StatusControls[$i].Foreground = '#00C896'
            }
            $uiSetup.Progress.Visibility    = 'Collapsed'
            $uiSetup.BtnInstall.IsEnabled   = $false
            $uiSetup.BtnContinue.Visibility = 'Visible'
            $uiSetup.Note.Text              = 'Installation complete. Click Continue to proceed.'
            $uiSetup.Note.Foreground        = '#00C896'
        })

    } catch {
        PLog "[ERROR] $_"
        $uiSetup.Window.Dispatcher.Invoke([Action]{
            $uiSetup.Progress.Visibility  = 'Collapsed'
            $uiSetup.BtnInstall.IsEnabled = $true
            $uiSetup.Note.Text            = 'Installation failed — see log above.'
            $uiSetup.Note.Foreground      = '#FF4757'
        })
    }
}

# ── Install button ────────────────────────────────────────────
$btnInstall.Add_Click({
    $btnInstall.IsEnabled      = $false
    $prereqProgress.Visibility = 'Visible'
    $prereqLog.Visibility      = 'Visible'
    $prereqNote.Text           = 'Installing, please wait…'
    $prereqNote.Foreground     = '#6060A0'

    $uiSetup = [hashtable]::Synchronized(@{
        Window         = $window
        Log            = $prereqLog
        Progress       = $prereqProgress
        BtnInstall     = $btnInstall
        BtnContinue    = $btnContinue
        Note           = $prereqNote
        DotControls    = $script:DotControls
        StatusControls = $script:StatusControls
    })

    $rs = [runspacefactory]::CreateRunspace()
    $rs.ApartmentState = "STA"; $rs.ThreadOptions = "ReuseThread"; $rs.Open()
    $ps = [powershell]::Create()
    $ps.Runspace = $rs
    [void]$ps.AddScript($prereqWorker)
    [void]$ps.AddArgument($uiSetup)
    [void]$ps.AddArgument($script:RequiredModules)
    $handle = $ps.BeginInvoke()

    $t = New-Object System.Windows.Threading.DispatcherTimer
    $t.Interval = [TimeSpan]::FromMilliseconds(400)
    $t.Add_Tick({
        if ($handle.IsCompleted) {
            $t.Stop()
            try { $ps.EndInvoke($handle) | Out-Null } catch {}
            $ps.Dispose(); $rs.Close(); $rs.Dispose()
        }
    })
    $t.Start()
})

# ── Continue button ───────────────────────────────────────────
$btnContinue.Add_Click({
    $setupOverlay.Visibility = 'Collapsed'
    $mainGrid.IsEnabled      = $true
    Import-Module Microsoft.Graph.Authentication -ErrorAction SilentlyContinue
})

# ── UI helpers (main thread) ──────────────────────────────────
function Write-UILog([string]$Msg) {
    $ts = Get-Date -Format "HH:mm:ss"
    $ui.Window.Dispatcher.Invoke([Action]{
        $ui.Log.AppendText("[$ts] $Msg`n")
        $ui.Log.ScrollToEnd()
    })
}

# ── Auth ──────────────────────────────────────────────────────
# Connect runs in a runspace so the UI thread is never blocked.
# For device code flow, Write-Host output (URL + code) is polled from
# the Information stream every 400 ms and forwarded to the log panel.
$btnConnect.Add_Click({
    $btnConnect.IsEnabled    = $false
    $rbDeviceCode.IsEnabled  = $false
    $rbInteractive.IsEnabled = $false
    $txtAuthStatus.Text      = "Connecting…"
    $txtAuthStatus.Foreground = '#9090C0'
    $useDeviceCode = [bool]$rbDeviceCode.IsChecked

    # Store connect-session objects in $ui (script-level) so the timer tick
    # can still reach them after this Add_Click handler returns
    $ui['ConnUi'] = [hashtable]::Synchronized(@{
        Window        = $window
        Log           = $txtLog
        AuthStatus    = $txtAuthStatus
        AuthDetail    = $txtAuthDetail
        BtnConnect    = $btnConnect
        BtnDisconnect = $btnDisconnect
        BtnRunH       = $btnRunHostname
        BtnRunS       = $btnRunSerial
        RbDevice      = $rbDeviceCode
        RbInteractive = $rbInteractive
        InfoIdx       = 0
    })

    $connectWorker = {
        param($cUI, $useDevCode)

        function L([string]$m) {
            $t = Get-Date -F "HH:mm:ss"
            $cUI.Window.Dispatcher.Invoke([Action]{
                $cUI.Log.AppendText("[$t] $m`n")
                $cUI.Log.ScrollToEnd()
            })
        }

        try {
            try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}
            Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

            $scopes = @("DeviceManagementManagedDevices.Read.All", "Device.Read.All")

            if ($useDevCode) {
                L "Starting authentication (device code flow)…"
                $InformationPreference = 'Continue'
                Connect-MgGraph -Scopes $scopes -UseDeviceAuthentication -NoWelcome -ErrorAction Stop 6>&1 |
                    ForEach-Object { $s = "$_".Trim(); if ($s) { L $s } }
            } else {
                L "Starting authentication (interactive browser)…"
                Connect-MgGraph -Scopes $scopes -NoWelcome -ErrorAction Stop
            }

            $ctx = Get-MgContext
            if (!$ctx) { throw "Get-MgContext returned null after connect." }

            $account  = $ctx.Account
            $tenantId = $ctx.TenantId
            $cUI.Window.Dispatcher.Invoke([Action]{
                $cUI.AuthStatus.Text          = "Connected  —  $account"
                $cUI.AuthStatus.Foreground    = '#00C896'
                $cUI.AuthDetail.Text          = "Tenant: $tenantId"
                $cUI.BtnDisconnect.Visibility = 'Visible'
                $cUI.BtnConnect.Content       = 'Reconnect'
                $cUI.BtnRunH.IsEnabled        = $true
                $cUI.BtnRunS.IsEnabled        = $true
            })
            L "Authenticated as $account  |  Tenant: $tenantId"

        } catch {
            $errMsg = "$_"
            L "[ERROR] Auth: $errMsg"
            $cUI.Window.Dispatcher.Invoke([Action]{
                $cUI.AuthStatus.Text       = "Connection failed"
                $cUI.AuthStatus.Foreground = '#FF4757'
                $cUI.AuthDetail.Text       = ""
            })

            if ($errMsg -match 'GetTokenAsync|does not have an implementation|TypeLoadException') {
                L "[HINT] Module version conflict — reinstall required, then restart this tool."
                $cUI.Window.Dispatcher.Invoke([Action]{
                    $fix = [System.Windows.MessageBox]::Show(
                        "Module version conflict detected:`n`n$errMsg`n`n" +
                        "Click YES to reinstall the modules automatically (requires admin).`n" +
                        "Click NO to fix manually:  Install-Module Microsoft.Graph -Force -Scope CurrentUser",
                        "Module Version Conflict","YesNo","Warning"
                    )
                    if ($fix -eq 'Yes') {
                        try {
                            $cUI.AuthStatus.Text = "Reinstalling module…"
                            $null = Start-Process powershell.exe -ArgumentList (
                                "-ExecutionPolicy Bypass -NoProfile -Command `"" +
                                "Get-InstalledModule -Name 'Microsoft.Graph*' -ErrorAction SilentlyContinue | Uninstall-Module -AllVersions -Force -ErrorAction SilentlyContinue; " +
                                "Install-Module Microsoft.Graph.Authentication -Force -Scope CurrentUser; " +
                                "Install-Module Microsoft.Graph.DeviceManagement -Force -Scope CurrentUser; " +
                                "Install-Module Microsoft.Graph.Identity.DirectoryManagement -Force -Scope CurrentUser`""
                            ) -WorkingDirectory 'C:\Windows\System32' -Verb RunAs -Wait -PassThru
                            [System.Windows.MessageBox]::Show(
                                "Module reinstall complete.`nPlease close and reopen this tool.",
                                "Done","OK","Information") | Out-Null
                        } catch {
                            [System.Windows.MessageBox]::Show(
                                "Reinstall failed:`n$_","Error","OK","Error") | Out-Null
                        }
                    }
                })
            } else {
                $cUI.Window.Dispatcher.Invoke([Action]{
                    [System.Windows.MessageBox]::Show(
                        "Authentication failed:`n`n$errMsg",
                        "Error","OK","Error") | Out-Null
                })
            }

        } finally {
            $cUI.Window.Dispatcher.Invoke([Action]{
                $cUI.BtnConnect.IsEnabled    = $true
                $cUI.RbDevice.IsEnabled      = $true
                $cUI.RbInteractive.IsEnabled = $true
            })
        }
    }

    $ui['ConnRs'] = [runspacefactory]::CreateRunspace()
    $ui['ConnRs'].ApartmentState = "STA"; $ui['ConnRs'].ThreadOptions = "ReuseThread"; $ui['ConnRs'].Open()
    $ui['ConnPs'] = [powershell]::Create()
    $ui['ConnPs'].Runspace = $ui['ConnRs']
    [void]$ui['ConnPs'].AddScript($connectWorker)
    [void]$ui['ConnPs'].AddArgument($ui['ConnUi'])
    [void]$ui['ConnPs'].AddArgument($useDeviceCode)
    $ui['ConnHandle'] = $ui['ConnPs'].BeginInvoke()

    $ui['ConnTimer'] = New-Object System.Windows.Threading.DispatcherTimer
    $ui['ConnTimer'].Interval = [TimeSpan]::FromMilliseconds(400)
    $ui['ConnTimer'].Add_Tick({
        $cUi = $ui['ConnUi']; $cPs = $ui['ConnPs']
        while ($cUi['InfoIdx'] -lt $cPs.Streams.Information.Count) {
            $msg = $cPs.Streams.Information[$cUi['InfoIdx']].MessageData
            if ($msg) {
                $ts = Get-Date -F "HH:mm:ss"
                $cUi.Log.AppendText("[$ts] $msg`n")
                $cUi.Log.ScrollToEnd()
            }
            $cUi['InfoIdx'] = [int]$cUi['InfoIdx'] + 1
        }
        if ($ui['ConnHandle'].IsCompleted) {
            $ui['ConnTimer'].Stop()
            try { $cPs.EndInvoke($ui['ConnHandle']) | Out-Null } catch {}
            $cPs.Dispose(); $ui['ConnRs'].Close(); $ui['ConnRs'].Dispose()
        }
    })
    $ui['ConnTimer'].Start()
})

$btnDisconnect.Add_Click({
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}
    $txtAuthStatus.Text       = "Not connected"
    $txtAuthStatus.Foreground = '#6060A0'
    $txtAuthDetail.Text       = ""
    $btnDisconnect.Visibility = 'Collapsed'
    $btnConnect.Content       = 'Connect to Microsoft Graph'
    $btnRunHostname.IsEnabled = $false
    $btnRunSerial.IsEnabled   = $false
    Write-UILog "Disconnected."
})

# ── Browse / Load ─────────────────────────────────────────────
$btnBrowseHostname.Add_Click({
    $d = New-Object System.Windows.Forms.OpenFileDialog
    $d.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
    if ($d.ShowDialog() -eq "OK") { $txtHostnameFile.Text = $d.FileName }
})
$btnLoadHostname.Add_Click({
    $p = $txtHostnameFile.Text.Trim()
    if (Test-Path $p) { $txtHostnameList.Text = (Get-Content $p -Raw); Write-UILog "Loaded: $p" }
    else { [System.Windows.MessageBox]::Show("File not found: $p","Error","OK","Error") | Out-Null }
})
$btnBrowseHostnameOut.Add_Click({
    $d = New-Object System.Windows.Forms.SaveFileDialog
    $d.Filter = "CSV (*.csv)|*.csv"; $d.FileName = "IntuneGroupImport.csv"
    if ($d.ShowDialog() -eq "OK") { $txtHostnameOutput.Text = $d.FileName }
})
$btnBrowseSerial.Add_Click({
    $d = New-Object System.Windows.Forms.OpenFileDialog
    $d.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
    if ($d.ShowDialog() -eq "OK") { $txtSerialFile.Text = $d.FileName }
})
$btnLoadSerial.Add_Click({
    $p = $txtSerialFile.Text.Trim()
    if (Test-Path $p) { $txtSerialList.Text = (Get-Content $p -Raw); Write-UILog "Loaded: $p" }
    else { [System.Windows.MessageBox]::Show("File not found: $p","Error","OK","Error") | Out-Null }
})
$btnBrowseSerialOut.Add_Click({
    $d = New-Object System.Windows.Forms.SaveFileDialog
    $d.Filter = "CSV (*.csv)|*.csv"; $d.FileName = "IntuneGroupImport.csv"
    if ($d.ShowDialog() -eq "OK") { $txtSerialOutput.Text = $d.FileName }
})
$btnClearLog.Add_Click({ $txtLog.Clear() })

# Hide placeholder when user types/pastes
$txtHostnameList.Add_TextChanged({ $phHostname.Visibility = if ($txtHostnameList.Text -eq '') {'Visible'} else {'Collapsed'} })
$txtSerialList.Add_TextChanged({   $phSerial.Visibility   = if ($txtSerialList.Text   -eq '') {'Visible'} else {'Collapsed'} })

# ── Template header lines (hardcoded) ─────────────────────────
$script:TemplateLines = @(
    "Member object ID or user principal name [memberObjectIdOrUpn] Required",
    "Example: 9832aad8-e4fe-496b-a604-95c6eF01ae75"
)

# ── Runspace launcher ─────────────────────────────────────────
function Start-GraphRunspace {
    param([scriptblock]$Worker, [object[]]$WorkerArgs)
    $rs = [runspacefactory]::CreateRunspace()
    $rs.ApartmentState = "STA"; $rs.ThreadOptions = "ReuseThread"; $rs.Open()
    $ps = [powershell]::Create()
    $ps.Runspace = $rs
    [void]$ps.AddScript($Worker)
    foreach ($a in $WorkerArgs) { [void]$ps.AddArgument($a) }
    $handle = $ps.BeginInvoke()
    $t = New-Object System.Windows.Threading.DispatcherTimer
    $t.Interval = [TimeSpan]::FromMilliseconds(400)
    $t.Add_Tick({
        if ($handle.IsCompleted) {
            $t.Stop()
            try { $ps.EndInvoke($handle) | Out-Null } catch { Write-UILog "[ERROR] $_" }
            $ps.Dispose(); $rs.Close(); $rs.Dispose()
        }
    })
    $t.Start()
}

# ── Worker: Hostname ──────────────────────────────────────────
$hostnameWorker = {
    param($ui, $hostnames, $outputPath, $templateLines)

    [System.Environment]::SetEnvironmentVariable("MSAL_DISABLE_WAM", "1", "Process")

    function L([string]$m) {
        $t = Get-Date -F "HH:mm:ss"
        $ui.Window.Dispatcher.Invoke([Action]{ $ui.Log.AppendText("[$t] $m`n"); $ui.Log.ScrollToEnd() })
    }
    function S([string]$m) { $ui.Window.Dispatcher.Invoke([Action]{ $ui.Status.Text = $m }) }
    function P([int]$c,[int]$n) {
        $p = if ($n) { [int](($c/$n)*100) } else { 0 }
        $ui.Window.Dispatcher.Invoke([Action]{
            $ui.Progress.IsIndeterminate = $false
            $ui.Progress.Value           = $p
            $ui.ProgressLabel.Text       = "$c / $n"
        })
    }

    try {
        Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
        Import-Module Microsoft.Graph.DeviceManagement -ErrorAction SilentlyContinue
        if (!(Get-MgContext)) { throw "Not authenticated. Connect first." }

        S "Loading managed devices from Intune…"
        $ui.Window.Dispatcher.Invoke([Action]{ $ui.Progress.IsIndeterminate = $true })
        L "Fetching all managed devices…"

        $all  = [System.Collections.Generic.List[object]]::new()
        $next = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$select=deviceName,id,azureADDeviceId"
        do {
            $r = Invoke-MgGraphRequest -Uri $next -Method GET
            foreach ($d in $r.value) { $all.Add($d) }
            $next = $r['@odata.nextLink']
        } while ($next)

        L "Loaded $($all.Count) devices."
        $total = $hostnames.Count
        $ids   = [System.Collections.Generic.List[string]]::new()
        $seen  = @{}
        $nf    = [System.Collections.Generic.List[string]]::new()
        $cur   = 0

        S "Processing $total hostnames…"
        foreach ($name in $hostnames) {
            $cur++; P $cur $total
            if ([string]::IsNullOrWhiteSpace($name)) { continue }
            $dev = $all | Where-Object { $_.deviceName -eq $name } | Select-Object -First 1
            if ($dev) {
                if (!$seen.ContainsKey($dev.id)) {
                    $seen[$dev.id] = $name
                    $ids.Add($dev.id)
                    L "FOUND    $name  ->  $($dev.id)"
                } else {
                    L "SKIP     $name  (duplicate Object ID)"
                }
            } else {
                $nf.Add($name)
                L "NOT FOUND  $name"
            }
        }

        $outDir = [System.IO.Path]::GetDirectoryName($outputPath)
        if (!(Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }
        $templateLines | Out-File -FilePath $outputPath -Encoding UTF8
        $ids           | Out-File -FilePath $outputPath -Append -Encoding UTF8

        $mapPath = ($outputPath -replace '\.csv$','') + "_mapping.csv"
        "Hostname,ObjectId,Status" | Out-File -FilePath $mapPath -Encoding UTF8
        foreach ($k in $seen.Keys)  { "$($seen[$k]),$k,Found"   | Out-File -FilePath $mapPath -Append -Encoding UTF8 }
        foreach ($n in $nf)         { "$n,,Not Found"            | Out-File -FilePath $mapPath -Append -Encoding UTF8 }

        $ui.Window.Dispatcher.Invoke([Action]{
            $ui.Progress.IsIndeterminate = $false
            $ui.Progress.Value           = 100
            $ui.ProgressLabel.Text       = "$($ids.Count) / $total found"
        })

        L "--- SUMMARY ---"
        L "Total: $total  |  Found: $($ids.Count)  |  Not found: $($nf.Count)"
        L "Output:  $outputPath"
        L "Mapping: $mapPath"
        S "Done — $($ids.Count) devices exported"

        [System.Windows.MessageBox]::Show(
            "Import complete!`n`nDevices found:  $($ids.Count)`nNot found:       $($nf.Count)`n`nOutput:`n$outputPath`n`nMapping:`n$mapPath",
            "Done","OK","Information"
        ) | Out-Null

    } catch {
        L "[ERROR] $_"; S "Error — see log"
        [System.Windows.MessageBox]::Show("Hostname import failed:`n`n$_","Error","OK","Error") | Out-Null
    } finally {
        $ui.Window.Dispatcher.Invoke([Action]{
            $ui.BtnRunH.IsEnabled = $true
            $ui.BtnRunS.IsEnabled = $true
            $ui.IsRunning         = $false
        })
    }
}

# ── Worker: Serial ────────────────────────────────────────────
$serialWorker = {
    param($ui, $serials, $outputPath, $templateLines)

    [System.Environment]::SetEnvironmentVariable("MSAL_DISABLE_WAM", "1", "Process")

    function L([string]$m) {
        $t = Get-Date -F "HH:mm:ss"
        $ui.Window.Dispatcher.Invoke([Action]{ $ui.Log.AppendText("[$t] $m`n"); $ui.Log.ScrollToEnd() })
    }
    function S([string]$m) { $ui.Window.Dispatcher.Invoke([Action]{ $ui.Status.Text = $m }) }
    function P([int]$c,[int]$n) {
        $p = if ($n) { [int](($c/$n)*100) } else { 0 }
        $ui.Window.Dispatcher.Invoke([Action]{
            $ui.Progress.IsIndeterminate = $false
            $ui.Progress.Value           = $p
            $ui.ProgressLabel.Text       = "$c / $n"
        })
    }

    try {
        Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
        Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction SilentlyContinue
        if (!(Get-MgContext)) { throw "Not authenticated. Connect first." }

        S "Loading devices from Entra ID…"
        $ui.Window.Dispatcher.Invoke([Action]{ $ui.Progress.IsIndeterminate = $true })
        L "Fetching all Entra devices…"

        $all  = [System.Collections.Generic.List[object]]::new()
        $next = "https://graph.microsoft.com/v1.0/devices?`$select=displayName,id"
        do {
            $r = Invoke-MgGraphRequest -Uri $next -Method GET
            foreach ($d in $r.value) { $all.Add($d) }
            $next = $r['@odata.nextLink']
        } while ($next)

        L "Loaded $($all.Count) devices."
        $total = $serials.Count
        $ids   = [System.Collections.Generic.List[string]]::new()
        $nf    = [System.Collections.Generic.List[string]]::new()
        $stats = @{ Exact=0; StartsWith=0; Contains=0; Fuzzy=0 }
        $cur   = 0

        S "Processing $total serial numbers…"
        foreach ($serial in $serials) {
            $cur++; P $cur $total
            if ([string]::IsNullOrWhiteSpace($serial)) { continue }

            $dev = $null; $method = $null
            $dev = $all | Where-Object { $_.displayName -eq $serial }                               | Select-Object -First 1; if ($dev) { $method = "Exact" }
            if (!$dev) { $dev = $all | Where-Object { $_.displayName -like "$serial*" }             | Select-Object -First 1; if ($dev) { $method = "StartsWith" } }
            if (!$dev) { $dev = $all | Where-Object { $_.displayName -like "*$serial*" }            | Select-Object -First 1; if ($dev) { $method = "Contains" } }
            if (!$dev) { $dev = $all | Where-Object { $_.displayName -match [regex]::Escape($serial) } | Select-Object -First 1; if ($dev) { $method = "Fuzzy" } }

            if ($dev) {
                $ids.Add($dev.id)
                $stats[$method]++
                L "[$method]  $serial  ->  $($dev.displayName)  ->  $($dev.id)"
            } else {
                $nf.Add($serial)
                L "NOT FOUND  $serial"
            }
        }

        $unique  = $ids | Select-Object -Unique
        $outDir  = [System.IO.Path]::GetDirectoryName($outputPath)
        if (!(Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }
        $templateLines | Out-File -FilePath $outputPath -Encoding UTF8
        $unique        | Out-File -FilePath $outputPath -Append -Encoding UTF8

        $ui.Window.Dispatcher.Invoke([Action]{
            $ui.Progress.IsIndeterminate = $false
            $ui.Progress.Value           = 100
            $ui.ProgressLabel.Text       = "$($unique.Count) / $total found"
        })

        L "--- SUMMARY ---"
        L "Total: $total  |  Found: $($ids.Count) (unique: $($unique.Count))  |  Not found: $($nf.Count)"
        L "  Exact: $($stats.Exact)  StartsWith: $($stats.StartsWith)  Contains: $($stats.Contains)  Fuzzy: $($stats.Fuzzy)"
        L "Output: $outputPath"
        S "Done — $($unique.Count) devices exported"

        [System.Windows.MessageBox]::Show(
            "Import complete!`n`nFound: $($ids.Count) (unique: $($unique.Count))`nNot found: $($nf.Count)`n`nExact: $($stats.Exact)  StartsWith: $($stats.StartsWith)  Contains: $($stats.Contains)  Fuzzy: $($stats.Fuzzy)`n`nOutput:`n$outputPath",
            "Done","OK","Information"
        ) | Out-Null

    } catch {
        L "[ERROR] $_"; S "Error — see log"
        [System.Windows.MessageBox]::Show("Serial import failed:`n`n$_","Error","OK","Error") | Out-Null
    } finally {
        $ui.Window.Dispatcher.Invoke([Action]{
            $ui.BtnRunH.IsEnabled = $true
            $ui.BtnRunS.IsEnabled = $true
            $ui.IsRunning         = $false
        })
    }
}

# ── Run buttons ───────────────────────────────────────────────
function Get-CleanList([string]$raw) {
    return ($raw -split "`r`n|`n|`r") |
        Where-Object { $_ -and $_ -notmatch "^\s*$" } |
        ForEach-Object { $_.Trim() } |
        Select-Object -Unique
}

function Invoke-Import {
    param([string]$RawList, [string]$Out, [scriptblock]$Worker, [string]$Type, [string]$Unit)
    if ($ui.IsRunning) { return }
    $list = Get-CleanList $RawList
    if (!$list) { [System.Windows.MessageBox]::Show("Enter or load a $Type list first.","No Input","OK","Warning") | Out-Null; return }
    if (!$Out)  { [System.Windows.MessageBox]::Show("Specify an output file path.","No Output","OK","Warning") | Out-Null; return }
    $ui.IsRunning             = $true
    $btnRunHostname.IsEnabled = $false
    $btnRunSerial.IsEnabled   = $false
    $txtLog.Clear()
    Write-UILog "Starting $Type import for $($list.Count) $Unit…"
    $txtStatus.Text = "Running…"
    Start-GraphRunspace -Worker $Worker -WorkerArgs @($ui, $list, $Out, $script:TemplateLines)
}

$btnRunHostname.Add_Click({ Invoke-Import $txtHostnameList.Text $txtHostnameOutput.Text.Trim() $hostnameWorker "hostname"      "devices" })
$btnRunSerial.Add_Click({   Invoke-Import $txtSerialList.Text   $txtSerialOutput.Text.Trim()   $serialWorker   "serial number" "serials" })

# ── Cleanup on close ──────────────────────────────────────────
$window.Add_Closed({
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}
})

# ── Show ──────────────────────────────────────────────────────
[void]$window.ShowDialog()
