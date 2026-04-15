#Requires -Version 5.1
$ErrorActionPreference='SilentlyContinue';Set-ExecutionPolicy Bypass -Scope Process -Force 2>$null
[System.Environment]::SetEnvironmentVariable("MSAL_DISABLE_WAM","1","Process");$env:MSAL_DISABLE_WAM=1
if([System.Threading.Thread]::CurrentThread.ApartmentState-ne'STA'){Start-Process powershell.exe -ArgumentList "-ExecutionPolicy Bypass -STA -NoProfile -File `"$PSCommandPath`"" -Wait;exit}
Add-Type -AssemblyName PresentationFramework,PresentationCore,WindowsBase,System.Windows.Forms

# ── XAML ─────────────────────────────────────────────────────
[xml]$Xaml = @'
<Window
  xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
  xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
  Title="Intune Bulk Group Importer" Height="760" Width="900" MinHeight="620" MinWidth="720"
  WindowStartupLocation="CenterScreen" Background="#09090B" FontFamily="Segoe UI" Foreground="#E4E4E7">

  <Window.Resources>
    <Style x:Key="Card" TargetType="Border">
      <Setter Property="Background" Value="#18181B"/>
      <Setter Property="BorderBrush" Value="#27272A"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="CornerRadius" Value="8"/>
    </Style>

    <Style TargetType="Button">
      <Setter Property="Foreground" Value="White"/><Setter Property="Padding" Value="14,8"/>
      <Setter Property="BorderThickness" Value="0"/><Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Background" Value="#27272A"/><Setter Property="FontWeight" Value="Medium"/>
      <Setter Property="Template"><Setter.Value><ControlTemplate TargetType="Button">
        <Border x:Name="bd" Background="{TemplateBinding Background}" CornerRadius="6" Padding="{TemplateBinding Padding}">
          <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
        </Border>
        <ControlTemplate.Triggers>
          <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="bd" Property="Opacity" Value="0.8"/></Trigger>
          <Trigger Property="IsEnabled" Value="False"><Setter TargetName="bd" Property="Opacity" Value="0.3"/></Trigger>
        </ControlTemplate.Triggers>
      </ControlTemplate></Setter.Value></Setter>
    </Style>

    <Style TargetType="TextBox">
      <Setter Property="Background" Value="#09090B"/><Setter Property="Foreground" Value="#E4E4E7"/>
      <Setter Property="BorderBrush" Value="#27272A"/><Setter Property="BorderThickness" Value="1"/>
      <Setter Property="CaretBrush" Value="#6366F1"/><Setter Property="FontFamily" Value="Consolas"/>
      <Setter Property="Padding" Value="10,8"/>
      <Setter Property="Template"><Setter.Value><ControlTemplate TargetType="TextBox">
        <Border x:Name="bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="1" CornerRadius="6">
          <ScrollViewer x:Name="PART_ContentHost" Margin="{TemplateBinding Padding}"/>
        </Border>
        <ControlTemplate.Triggers>
          <Trigger Property="IsFocused" Value="True"><Setter TargetName="bd" Property="BorderBrush" Value="#6366F1"/></Trigger>
        </ControlTemplate.Triggers>
      </ControlTemplate></Setter.Value></Setter>
    </Style>
  </Window.Resources>

  <Grid x:Name="mainGrid" Margin="24" IsEnabled="False">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="180"/>
    </Grid.RowDefinitions>

    <StackPanel Grid.Row="0" Margin="0,0,0,16">
      <TextBlock Text="Intune Bulk Group Importer" FontSize="22" FontWeight="Bold" Foreground="#FFFFFF"/>
      <TextBlock Text="Resolve devices via Microsoft Graph and generate CSV" FontSize="12" Foreground="#A1A1AA" Margin="0,4,0,16"/>
      
      <Border Style="{StaticResource Card}" Padding="16,12">
        <DockPanel>
          <Button x:Name="btnConnect" DockPanel.Dock="Right" Content="Connect to Graph" Background="#6366F1" FontWeight="Bold"/>
          <Button x:Name="btnDisconnect" DockPanel.Dock="Right" Content="Disconnect" Background="#BE123C" Margin="0,0,10,0" Visibility="Collapsed"/>
          <StackPanel VerticalAlignment="Center">
            <TextBlock x:Name="txtAuthStatus" Text="Not connected" FontSize="13" FontWeight="SemiBold" Foreground="#A1A1AA"/>
            <TextBlock x:Name="txtAuthDetail" Text="Authentication: Interactive Browser" FontSize="11" Foreground="#71717A" Margin="0,4,0,0"/>
          </StackPanel>
        </DockPanel>
      </Border>
    </StackPanel>

    <TabControl x:Name="tabMain" Grid.Row="1" Background="#18181B" BorderBrush="#27272A" BorderThickness="1" Margin="0,0,0,16">
      <TabControl.Resources><Style TargetType="TabItem">
        <Setter Property="Background" Value="Transparent"/><Setter Property="Foreground" Value="#A1A1AA"/>
        <Setter Property="FontSize" Value="13"/><Setter Property="Padding" Value="20,10"/><Setter Property="Cursor" Value="Hand"/>
        <Setter Property="Template"><Setter.Value><ControlTemplate TargetType="TabItem">
          <Border x:Name="bd" Background="Transparent" BorderThickness="0,0,0,2" BorderBrush="Transparent" Padding="{TemplateBinding Padding}">
            <TextBlock Text="{TemplateBinding Header}" Foreground="{TemplateBinding Foreground}" FontWeight="SemiBold"/>
          </Border>
          <ControlTemplate.Triggers>
            <Trigger Property="IsSelected" Value="True"><Setter TargetName="bd" Property="BorderBrush" Value="#6366F1"/><Setter Property="Foreground" Value="#FFFFFF"/></Trigger>
            <Trigger Property="IsMouseOver" Value="True"><Setter Property="Foreground" Value="#E4E4E7"/></Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate></Setter.Value></Setter>
      </Style></TabControl.Resources>

      <TabItem Header="By Hostname">
        <Grid Margin="16">
          <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
          <DockPanel Grid.Row="0" Margin="0,0,0,12">
            <Button x:Name="btnLoadHostname" DockPanel.Dock="Right" Content="Load" Margin="8,0,0,0"/>
            <Button x:Name="btnBrowseHostname" DockPanel.Dock="Right" Content="Browse…" Margin="8,0,0,0"/>
            <TextBox x:Name="txtHostnameFile" Text="C:\TEMP\pc.txt"/>
          </DockPanel>
          <Grid Grid.Row="1">
            <TextBox x:Name="txtHostnameList" AcceptsReturn="True" TextWrapping="NoWrap" VerticalScrollBarVisibility="Auto"/>
            <TextBlock x:Name="phHostname" Text="Paste hostnames here, one per line&#x0a;Example: LAPTOP-001" Foreground="#52525B" FontFamily="Consolas" Margin="12,10" IsHitTestVisible="False"/>
          </Grid>
          <DockPanel Grid.Row="2" Margin="0,12,0,0">
            <Button x:Name="btnRunHostname" DockPanel.Dock="Right" Content=" Run Import " Background="#6366F1" IsEnabled="False"/>
            <Button x:Name="btnBrowseHostnameOut" DockPanel.Dock="Right" Content="…" Width="36" Margin="8,0"/>
            <TextBlock Text="Output:" Foreground="#A1A1AA" VerticalAlignment="Center" Margin="0,0,10,0"/>
            <TextBox x:Name="txtHostnameOutput" Text="C:\TEMP\IntuneGroupImport.csv"/>
          </DockPanel>
        </Grid>
      </TabItem>

      <TabItem Header="By Serial Number">
        <Grid Margin="16">
          <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
          <DockPanel Grid.Row="0" Margin="0,0,0,12">
            <Button x:Name="btnLoadSerial" DockPanel.Dock="Right" Content="Load" Margin="8,0,0,0"/>
            <Button x:Name="btnBrowseSerial" DockPanel.Dock="Right" Content="Browse…" Margin="8,0,0,0"/>
            <TextBox x:Name="txtSerialFile" Text="C:\TEMP\serials.txt"/>
          </DockPanel>
          <Grid Grid.Row="1">
            <TextBox x:Name="txtSerialList" AcceptsReturn="True" TextWrapping="NoWrap" VerticalScrollBarVisibility="Auto"/>
            <TextBlock x:Name="phSerial" Text="Paste serials here, one per line&#x0a;Example: C02XG2JHJTD5" Foreground="#52525B" FontFamily="Consolas" Margin="12,10" IsHitTestVisible="False"/>
          </Grid>
          <DockPanel Grid.Row="2" Margin="0,12,0,0">
            <Button x:Name="btnRunSerial" DockPanel.Dock="Right" Content=" Run Import " Background="#6366F1" IsEnabled="False"/>
            <Button x:Name="btnBrowseSerialOut" DockPanel.Dock="Right" Content="…" Width="36" Margin="8,0"/>
            <TextBlock Text="Output:" Foreground="#A1A1AA" VerticalAlignment="Center" Margin="0,0,10,0"/>
            <TextBox x:Name="txtSerialOutput" Text="C:\TEMP\IntuneGroupImport.csv"/>
          </DockPanel>
        </Grid>
      </TabItem>
    </TabControl>

    <StackPanel Grid.Row="2" Margin="0,0,0,12">
      <DockPanel>
        <TextBlock x:Name="txtProgressLabel" DockPanel.Dock="Right" Foreground="#6366F1" FontSize="11" FontWeight="Bold"/>
        <TextBlock x:Name="txtStatus" Text="Ready" Foreground="#A1A1AA" FontSize="11"/>
      </DockPanel>
      <ProgressBar x:Name="progressBar" Height="4" Margin="0,6,0,0" Background="#18181B" Foreground="#6366F1" BorderThickness="0" Maximum="100"/>
    </StackPanel>

    <Border Grid.Row="3" Style="{StaticResource Card}" Padding="0">
      <DockPanel>
        <DockPanel DockPanel.Dock="Top" Margin="12,8,12,6">
          <Button x:Name="btnClearLog" DockPanel.Dock="Right" Content="Clear" Padding="10,4" FontSize="10"/>
          <TextBlock Text="ACTIVITY LOG" Foreground="#52525B" FontSize="10" FontWeight="Bold" VerticalAlignment="Center"/>
        </DockPanel>
        <RichTextBox x:Name="txtLog" Background="Transparent" Foreground="#A1A1AA" BorderThickness="0" FontFamily="Consolas" FontSize="11" IsReadOnly="True" VerticalScrollBarVisibility="Auto" Padding="12,0,12,10"/>
      </DockPanel>
    </Border>

    <Border x:Name="setupOverlay" Background="#E6000000" Panel.ZIndex="100" Visibility="Collapsed">
      <Border Style="{StaticResource Card}" Padding="32" Width="520" VerticalAlignment="Center">
        <StackPanel>
          <TextBlock Text="Prerequisites Required" FontSize="20" FontWeight="Bold" Foreground="#FFFFFF" Margin="0,0,0,6"/>
          <TextBlock Text="Missing PowerShell modules. Click Install to set them up automatically." FontSize="13" Foreground="#A1A1AA" TextWrapping="Wrap" Margin="0,0,0,20"/>
          
          <Border Background="#09090B" BorderBrush="#27272A" BorderThickness="1" CornerRadius="6" Padding="20,16">
            <StackPanel>
              <DockPanel Margin="0,0,0,12"><TextBlock x:Name="modStatus0" DockPanel.Dock="Right" Text="—" Foreground="#71717A"/><Ellipse x:Name="dot0" Width="8" Height="8" Fill="#3F3F46" Margin="0,0,12,0"/><TextBlock Text="Microsoft.Graph.Authentication" FontFamily="Consolas"/></DockPanel>
              <DockPanel Margin="0,0,0,12"><TextBlock x:Name="modStatus1" DockPanel.Dock="Right" Text="—" Foreground="#71717A"/><Ellipse x:Name="dot1" Width="8" Height="8" Fill="#3F3F46" Margin="0,0,12,0"/><TextBlock Text="Microsoft.Graph.DeviceManagement" FontFamily="Consolas"/></DockPanel>
              <DockPanel><TextBlock x:Name="modStatus2" DockPanel.Dock="Right" Text="—" Foreground="#71717A"/><Ellipse x:Name="dot2" Width="8" Height="8" Fill="#3F3F46" Margin="0,0,12,0"/><TextBlock Text="Microsoft.Graph.Identity.DirectoryManagement" FontFamily="Consolas"/></DockPanel>
            </StackPanel>
          </Border>

          <ProgressBar x:Name="prereqProgress" Height="4" Margin="0,20,0,0" Background="#09090B" Foreground="#6366F1" BorderThickness="0" IsIndeterminate="True" Visibility="Collapsed"/>
          <TextBox x:Name="prereqLog" Height="100" Margin="0,12,0,0" IsReadOnly="True" VerticalScrollBarVisibility="Auto" Visibility="Collapsed"/>

          <DockPanel Margin="0,24,0,0">
            <Button x:Name="btnContinue" DockPanel.Dock="Right" Content="Continue →" Background="#10B981" Visibility="Collapsed"/>
            <Button x:Name="btnInstall" DockPanel.Dock="Right" Content="Install Modules" Background="#6366F1" Margin="10,0,0,0"/>
            <TextBlock x:Name="prereqNote" Text="Installed for current user only." Foreground="#71717A" FontSize="11" VerticalAlignment="Center"/>
          </DockPanel>
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
@('btnInstall','btnContinue','txtAuthStatus','txtAuthDetail','btnConnect','btnDisconnect','rbDeviceCode','rbInteractive','txtHostnameFile','btnBrowseHostname','btnLoadHostname','txtHostnameList','txtHostnameOutput','btnBrowseHostnameOut','btnRunHostname','phHostname','phSerial','txtSerialFile','btnBrowseSerial','btnLoadSerial','txtSerialList','txtSerialOutput','btnBrowseSerialOut','btnRunSerial','txtStatus','txtProgressLabel','progressBar','txtLog','btnClearLog')|%{Set-Variable $_ $window.FindName($_)}
$modStatus0=$window.FindName("modStatus0");$modStatus1=$window.FindName("modStatus1");$modStatus2=$window.FindName("modStatus2");$prereqProgress=$window.FindName("prereqProgress");$prereqLog=$window.FindName("prereqLog");$prereqNote=$window.FindName("prereqNote");$setupOverlay=$window.FindName("setupOverlay");$mainGrid=$window.FindName("mainGrid");$dot0=$window.FindName("dot0");$dot1=$window.FindName("dot1");$dot2=$window.FindName("dot2")

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
$script:RequiredModules=@("Microsoft.Graph.Authentication","Microsoft.Graph.DeviceManagement","Microsoft.Graph.Identity.DirectoryManagement")
$script:DotControls=@($dot0,$dot1,$dot2);$script:StatusControls=@($modStatus0,$modStatus1,$modStatus2)
function Test-Prerequisites{$allOk=$true;for($i=0;$i-lt$script:RequiredModules.Count;$i++){$ok=[bool](Get-Module -ListAvailable -Name $script:RequiredModules[$i]);$color=if($ok){'#00C896'}else{'#FF4757'};$script:DotControls[$i].Fill=$color;$script:StatusControls[$i].Text=if($ok){'Installed'}else{'Missing'};$script:StatusControls[$i].Foreground=$color;if(!$ok){$allOk=$false}};return $allOk}
$window.Add_Loaded({$allOk=Test-Prerequisites;if($allOk){$mainGrid.IsEnabled=$true;Import-Module Microsoft.Graph.Authentication -ErrorAction SilentlyContinue}else{$setupOverlay.Visibility='Visible'}})

# ── Install prerequisites worker ──────────────────────────────
$prereqWorker = {
    param($uiSetup, $modules)

    function PLog([string]$m) {
        $t=Get-Date -F "HH:mm:ss";$c=if($m-match"\[ERROR\]|failed"){"#FF6666"}elseif($m-match"Installed|OK"){"#00C896"}else{"#8080B8"}
        $uiSetup.Window.Dispatcher.Invoke([Action]{$p=New-Object System.Windows.Documents.Paragraph;$p.Margin=0;$tr=New-Object System.Windows.Documents.Run("[$t] ");$tr.Foreground="#7575A5";$p.Inlines.Add($tr);$mr=New-Object System.Windows.Documents.Run($m);$mr.Foreground=$c;$p.Inlines.Add($mr);$uiSetup.Log.Document.Blocks.Add($p);$uiSetup.Log.ScrollToEnd()})
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
function Write-UILog([string]$Msg){$ts=Get-Date -F "HH:mm:ss";$c=if($Msg-match"\[ERROR\]"){"#FF6666"}elseif($Msg-match"\[WARNING\]|security|Security|SECURITY"){"#F4B860"}elseif($Msg-match"SUCCESS|\[OK\]|Authenticated|Connected"){"#00C896"}else{"#8080B8"};$ui.Window.Dispatcher.Invoke([Action]{$p=New-Object System.Windows.Documents.Paragraph;$p.Margin=0;$tr=New-Object System.Windows.Documents.Run("[$ts] ");$tr.Foreground="#7575A5";$p.Inlines.Add($tr);$mr=New-Object System.Windows.Documents.Run($Msg);$mr.Foreground=$c;$p.Inlines.Add($mr);$ui.Log.Document.Blocks.Add($p);$ui.Log.ScrollToEnd()})}

# ── Auth ──────────────────────────────────────────────────────
# Connect runs in a runspace so the UI thread is never blocked.
# For device code flow, Write-Host output (URL + code) is polled from
# the Information stream every 400 ms and forwarded to the log panel.
$btnConnect.Add_Click({
    $btnConnect.IsEnabled=$false
    $txtAuthStatus.Text="Connecting…"
    $txtAuthStatus.Foreground='#9090C0'

    # Store connect-session objects in $ui (script-level) so the timer tick
    # can still reach them after this Add_Click handler returns
    $ui['ConnUi']=[hashtable]::Synchronized(@{
        Window=$window;Log=$txtLog;AuthStatus=$txtAuthStatus;AuthDetail=$txtAuthDetail;BtnConnect=$btnConnect;BtnDisconnect=$btnDisconnect
        BtnRunH=$btnRunHostname;BtnRunS=$btnRunSerial;InfoIdx=0
    })

    $connectWorker = {
        param($cUI)

        function L([string]$m){$t=Get-Date -F "HH:mm:ss";$c=if($m-match"\[ERROR\]"){"#FF6666"}elseif($m-match"\[WARNING\]|security|Security|SECURITY"){"#F4B860"}elseif($m-match"SUCCESS|\[OK\]|Authenticated|Connected"){"#00C896"}else{"#8080B8"};$cUI.Window.Dispatcher.Invoke([Action]{$p=New-Object System.Windows.Documents.Paragraph;$p.Margin=0;$tr=New-Object System.Windows.Documents.Run("[$t] ");$tr.Foreground="#7575A5";$p.Inlines.Add($tr);$mr=New-Object System.Windows.Documents.Run($m);$mr.Foreground=$c;$p.Inlines.Add($mr);$cUI.Log.Document.Blocks.Add($p);$cUI.Log.ScrollToEnd()})}

        try {
            # Disable WAM in this runspace to avoid window handle errors
            $env:MSAL_DISABLE_WAM = 1

            try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}

            # Import all required modules
            Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
            Import-Module Microsoft.Graph.DeviceManagement -ErrorAction Stop
            Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop

            $scopes=@("DeviceManagementManagedDevices.Read.All")
            L "Starting authentication (interactive browser)…"
            Connect-MgGraph -Scopes $scopes -NoWelcome -ErrorAction Stop

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
                $cUI.BtnConnect.IsEnabled=$true
            })
        }
    }

    $ui['ConnRs'] = [runspacefactory]::CreateRunspace()
    $ui['ConnRs'].ApartmentState = "STA"; $ui['ConnRs'].ThreadOptions = "ReuseThread"; $ui['ConnRs'].Open()
    $ui['ConnPs'] = [powershell]::Create()
    $ui['ConnPs'].Runspace=$ui['ConnRs']
    [void]$ui['ConnPs'].AddScript($connectWorker)
    [void]$ui['ConnPs'].AddArgument($ui['ConnUi'])
    $ui['ConnHandle']=$ui['ConnPs'].BeginInvoke()

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

$btnDisconnect.Add_Click({try{Disconnect-MgGraph -ErrorAction SilentlyContinue}catch{};$txtAuthStatus.Text="Not connected";$txtAuthStatus.Foreground='#6060A0';$txtAuthDetail.Text="";$btnDisconnect.Visibility='Collapsed';$btnConnect.Content='Connect to Microsoft Graph';$btnRunHostname.IsEnabled=$false;$btnRunSerial.IsEnabled=$false;Write-UILog"Disconnected."})

$btnBrowseHostname.Add_Click({$d=New-Object System.Windows.Forms.OpenFileDialog;$d.Filter="Text files (*.txt)|*.txt|All files (*.*)|*.*";if($d.ShowDialog()-eq"OK"){$txtHostnameFile.Text=$d.FileName}})
$btnLoadHostname.Add_Click({$p=$txtHostnameFile.Text.Trim();if(Test-Path $p){$txtHostnameList.Text=(Get-Content $p -Raw);Write-UILog"Loaded: $p"}else{[System.Windows.MessageBox]::Show("File not found: $p","Error","OK","Error")|Out-Null}})
$btnBrowseHostnameOut.Add_Click({$d=New-Object System.Windows.Forms.SaveFileDialog;$d.Filter="CSV (*.csv)|*.csv";$d.FileName="IntuneGroupImport.csv";if($d.ShowDialog()-eq"OK"){$txtHostnameOutput.Text=$d.FileName}})
$btnBrowseSerial.Add_Click({$d=New-Object System.Windows.Forms.OpenFileDialog;$d.Filter="Text files (*.txt)|*.txt|All files (*.*)|*.*";if($d.ShowDialog()-eq"OK"){$txtSerialFile.Text=$d.FileName}})
$btnLoadSerial.Add_Click({$p=$txtSerialFile.Text.Trim();if(Test-Path $p){$txtSerialList.Text=(Get-Content $p -Raw);Write-UILog"Loaded: $p"}else{[System.Windows.MessageBox]::Show("File not found: $p","Error","OK","Error")|Out-Null}})
$btnBrowseSerialOut.Add_Click({$d=New-Object System.Windows.Forms.SaveFileDialog;$d.Filter="CSV (*.csv)|*.csv";$d.FileName="IntuneGroupImport.csv";if($d.ShowDialog()-eq"OK"){$txtSerialOutput.Text=$d.FileName}})
$btnClearLog.Add_Click({$txtLog.Clear()})

$txtHostnameList.Add_TextChanged({$phHostname.Visibility=if($txtHostnameList.Text-eq''){'Visible'}else{'Collapsed'}})
$txtSerialList.Add_TextChanged({$phSerial.Visibility=if($txtSerialList.Text-eq''){'Visible'}else{'Collapsed'}})
$script:TemplateLines=@("Member object ID or user principal name [memberObjectIdOrUpn] Required","Example: 9832aad8-e4fe-496b-a604-95c6eF01ae75")

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
        $color = "#8080B8"
        if ($m -match "\[ERROR\]|NOT FOUND") { $color = "#FF6666" }
        elseif ($m -match "\[WARNING\]|SKIP") { $color = "#F4B860" }
        elseif ($m -match "FOUND|SUCCESS") { $color = "#00C896" }

        $ui.Window.Dispatcher.Invoke([Action]{
            $para = New-Object System.Windows.Documents.Paragraph
            $para.Margin = New-Object System.Windows.Thickness(0,0,0,0)
            $tsRun = New-Object System.Windows.Documents.Run("[$t] ")
            $tsRun.Foreground = "#7575A5"
            $para.Inlines.Add($tsRun)
            $msgRun = New-Object System.Windows.Documents.Run($m)
            $msgRun.Foreground = $color
            $para.Inlines.Add($msgRun)
            $ui.Log.Document.Blocks.Add($para)
            $ui.Log.ScrollToEnd()
        })
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
        $color = "#8080B8"
        if ($m -match "\[ERROR\]|NOT FOUND") { $color = "#FF6666" }
        elseif ($m -match "\[WARNING\]|SKIP") { $color = "#F4B860" }
        elseif ($m -match "FOUND|SUCCESS") { $color = "#00C896" }

        $ui.Window.Dispatcher.Invoke([Action]{
            $para = New-Object System.Windows.Documents.Paragraph
            $para.Margin = New-Object System.Windows.Thickness(0,0,0,0)
            $tsRun = New-Object System.Windows.Documents.Run("[$t] ")
            $tsRun.Foreground = "#7575A5"
            $para.Inlines.Add($tsRun)
            $msgRun = New-Object System.Windows.Documents.Run($m)
            $msgRun.Foreground = $color
            $para.Inlines.Add($msgRun)
            $ui.Log.Document.Blocks.Add($para)
            $ui.Log.ScrollToEnd()
        })
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
