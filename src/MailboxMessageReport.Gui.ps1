<#
.SYNOPSIS
    Mailbox Message Report - window (dot-sourced by MailboxMessageReport.psm1).

.DESCRIPTION
    WPF window. With .NET 9 or later (PowerShell 7.5 and later) it uses the Fluent theme of Windows 11: light or dark
    like Windows, rounded controls, with the accent colour of the report (#B11F4B). With PowerShell 7.4 (.NET 8) the
    same window uses the classic WPF controls with the colours of the report.

    The window is made to type the search, follow it and open the report: a mailbox can hold hundreds of thousands
    of messages, the report is the place to read them. Layout: a header; on the left the mailboxes, the period and the
    subjects, where to read (primary mailbox, archive, Recoverable Items), the report and the connection; on the right
    the mailboxes read (counts per location), a preview of the first messages of the report (Window.PreviewMessages),
    and the progress; at the bottom the buttons.

    It runs exactly the same engine as the command line (Find-MmrMessages, Export-MmrReport): the progress shows the
    lines of the console. The work runs in a runspace of its own (the module loaded there when the window opens): the
    window always answers, its lines come through a queue read every 100 ms (Start-MmrGuiWork, Step-MmrGuiWork); Stop
    ends the run at the next page. The progress bar (step, part done, time left) and the taskbar button follow the run.

    Closing never needs PowerShell code: Close is the cancel button of the window (Esc too) and the title-bar button is
    native; the only Closing handler is attached while a run is in progress. Ctrl+C is ignored in the console while the
    window is open.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.0.0
#>

$script:Gui = $null

function Get-MmrGuiXaml {
    <# The window. Colours come from the Fluent theme resources (or the classic fallback of Set-MmrGuiTheme). #>
    @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Width="1320" Height="900" MinWidth="1080" MinHeight="680" WindowStartupLocation="CenterScreen"
        FontFamily="Segoe UI Variable Text, Segoe UI" FontSize="14" UseLayoutRounding="True">
  <Window.Resources>
    <Style x:Key="MmrCard" TargetType="Border">
      <Setter Property="Background" Value="{DynamicResource CardBackgroundFillColorDefaultBrush}"/>
      <Setter Property="BorderBrush" Value="{DynamicResource CardStrokeColorDefaultBrush}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="CornerRadius" Value="8"/>
      <Setter Property="Padding" Value="18,14,18,16"/>
      <Setter Property="Margin" Value="0,0,0,12"/>
    </Style>
    <Style x:Key="MmrCardTitle" TargetType="TextBlock">
      <Setter Property="FontSize" Value="14"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Margin" Value="0,0,0,8"/>
      <Setter Property="Foreground" Value="{DynamicResource TextFillColorPrimaryBrush}"/>
    </Style>
    <Style x:Key="MmrLabel" TargetType="TextBlock">
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="Foreground" Value="{DynamicResource TextFillColorSecondaryBrush}"/>
      <Setter Property="Margin" Value="0,8,0,4"/>
    </Style>
    <Style x:Key="MmrHint" TargetType="TextBlock">
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="TextWrapping" Value="Wrap"/>
      <Setter Property="Foreground" Value="{DynamicResource TextFillColorSecondaryBrush}"/>
    </Style>
    <Style x:Key="MmrIcon" TargetType="TextBlock">
      <Setter Property="FontFamily" Value="Segoe Fluent Icons, Segoe MDL2 Assets"/>
    </Style>
    <Style x:Key="MmrRow" TargetType="ListViewItem" BasedOn="{StaticResource {x:Static GridView.GridViewItemContainerStyleKey}}">
      <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
      <Setter Property="Padding" Value="0,1"/>
      <Setter Property="MinHeight" Value="28"/>
    </Style>
  </Window.Resources>

  <Grid x:Name="Root" Background="{DynamicResource ApplicationBackgroundBrush}">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <Grid x:Name="Header" Margin="24,18,24,14">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="Auto"/>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="Auto"/>
      </Grid.ColumnDefinitions>
      <Border Width="46" Height="46" CornerRadius="10" Background="{DynamicResource MmrBrand}" VerticalAlignment="Center">
        <TextBlock Style="{StaticResource MmrIcon}" Text="&#xE715;" FontSize="22" Foreground="White" HorizontalAlignment="Center" VerticalAlignment="Center"/>
      </Border>
      <StackPanel Grid.Column="1" Margin="14,0,0,0" VerticalAlignment="Center">
        <TextBlock Text="EXCHANGE ONLINE MAILBOX AND ARCHIVE" FontSize="11" FontWeight="SemiBold" Foreground="{DynamicResource MmrBrandText}"/>
        <TextBlock Text="Mailbox Message Report" FontSize="24" FontWeight="SemiBold" Foreground="{DynamicResource TextFillColorPrimaryBrush}"/>
        <TextBlock FontSize="13" Foreground="{DynamicResource TextFillColorSecondaryBrush}" TextTrimming="CharacterEllipsis"
                   Text="List the messages of the primary mailbox and of the archive: period, subjects, folders, sender and recipients. Read only."/>
      </StackPanel>
      <TextBlock x:Name="Version" Grid.Column="2" FontSize="12" Foreground="{DynamicResource TextFillColorSecondaryBrush}" VerticalAlignment="Top"/>
    </Grid>

    <Grid Grid.Row="1" Margin="24,0,24,0">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="410"/>
        <ColumnDefinition Width="16"/>
        <ColumnDefinition Width="*"/>
      </Grid.ColumnDefinitions>
      <ScrollViewer x:Name="SettingsScroll" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled" Padding="0,0,4,0">
        <StackPanel x:Name="Inputs">
          <Border Style="{StaticResource MmrCard}">
            <StackPanel>
              <Grid>
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock Text="Mailboxes" Style="{StaticResource MmrCardTitle}"/>
                <Button x:Name="LoadMailboxes" Grid.Column="1" Content="Load a list..." Padding="10,2" Margin="0,-4,0,6"/>
              </Grid>
              <TextBox x:Name="Mailbox" AcceptsReturn="True" TextWrapping="NoWrap" MinHeight="34" MaxHeight="110" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"/>
              <TextBlock x:Name="MailboxHint" Style="{StaticResource MmrHint}" Margin="0,6,0,0"/>
            </StackPanel>
          </Border>

          <Border Style="{StaticResource MmrCard}">
            <StackPanel>
              <TextBlock Text="Messages" Style="{StaticResource MmrCardTitle}" Margin="0,0,0,0"/>
              <Grid>
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="12"/>
                  <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <StackPanel>
                  <TextBlock Text="Received from" Style="{StaticResource MmrLabel}"/>
                  <DatePicker x:Name="StartDate" SelectedDateFormat="Short"/>
                </StackPanel>
                <StackPanel Grid.Column="2">
                  <TextBlock Text="To (included)" Style="{StaticResource MmrLabel}"/>
                  <DatePicker x:Name="EndDate" SelectedDateFormat="Short"/>
                </StackPanel>
              </Grid>
              <TextBlock Style="{StaticResource MmrHint}" Margin="0,6,0,0" Text="Empty: no limit. The received date of the message."/>
              <TextBlock Text="Subject contains (one per line: any of them; empty = every message)" Style="{StaticResource MmrLabel}"/>
              <TextBox x:Name="Subject" AcceptsReturn="True" TextWrapping="NoWrap" MinHeight="34" MaxHeight="90" VerticalScrollBarVisibility="Auto"/>
              <CheckBox x:Name="Recipients" Content="Recipients (To, Cc, Bcc)" Margin="0,10,0,0"/>
              <TextBlock Style="{StaticResource MmrHint}" Margin="28,0,0,0" Text="Exchange reads the recipients of each message for them: untick for large folders of meeting messages (much faster)."/>
            </StackPanel>
          </Border>

          <Border Style="{StaticResource MmrCard}">
            <StackPanel>
              <TextBlock Text="Where" Style="{StaticResource MmrCardTitle}"/>
              <CheckBox x:Name="LocationPrimary" Content="Primary mailbox"/>
              <TextBlock Style="{StaticResource MmrHint}" Margin="28,0,0,6" Text="Every mail folder: Inbox, Sent Items, Deleted Items, the folders of the user."/>
              <CheckBox x:Name="LocationArchive" Content="Archive (In-Place Archive)"/>
              <TextBlock Style="{StaticResource MmrHint}" Margin="28,0,0,6" Text="Read through Microsoft Graph as a mailbox of its own (MBX:&lt;ArchiveGuid&gt;). Not the auxiliary archives of an auto-expanding archive."/>
              <CheckBox x:Name="RecoverableItems" Content="Recoverable Items"/>
              <TextBlock Style="{StaticResource MmrHint}" Margin="28,0,0,6" Text="Deletions, Purges, Versions, DiscoveryHolds...: the items deleted from Deleted Items, or kept by a hold."/>
              <TextBlock Text="Folders left out (one path per line, * allowed)" Style="{StaticResource MmrLabel}"/>
              <TextBox x:Name="ExcludeFolders" AcceptsReturn="True" TextWrapping="NoWrap" MinHeight="34" MaxHeight="80" VerticalScrollBarVisibility="Auto"/>
            </StackPanel>
          </Border>

          <Border Style="{StaticResource MmrCard}">
            <StackPanel>
              <TextBlock Text="Report" Style="{StaticResource MmrCardTitle}"/>
              <StackPanel Orientation="Horizontal" Margin="0,0,0,8">
                <CheckBox x:Name="FormatCsv" Content="CSV" Margin="0,0,24,0"/>
                <CheckBox x:Name="FormatHtml" Content="HTML"/>
              </StackPanel>
              <RadioButton x:Name="LayoutGlobal" GroupName="Layout" Content="One report for every mailbox"/>
              <RadioButton x:Name="LayoutPerMailbox" GroupName="Layout" Content="One report per mailbox"/>
              <RadioButton x:Name="LayoutBoth" GroupName="Layout" Content="Both"/>
              <TextBlock x:Name="ReportHint" Style="{StaticResource MmrHint}" Margin="0,6,0,0"/>
            </StackPanel>
          </Border>

          <Expander x:Name="ConnectionExpander" Header="Connection (Microsoft Graph application)" Margin="0,0,0,12">
            <Border Style="{StaticResource MmrCard}" Margin="0,8,0,0">
              <StackPanel>
                <TextBlock Text="Tenant ID or domain" Style="{StaticResource MmrLabel}" Margin="0,0,0,4"/>
                <TextBox x:Name="TenantId"/>
                <TextBlock Text="Application (client) ID" Style="{StaticResource MmrLabel}"/>
                <TextBox x:Name="AppId"/>
                <TextBlock Text="Sign-in of the application" Style="{StaticResource MmrLabel}"/>
                <ComboBox x:Name="AuthMode"/>
                <StackPanel x:Name="ThumbPanel">
                  <TextBlock Text="Certificate thumbprint" Style="{StaticResource MmrLabel}"/>
                  <TextBox x:Name="Thumbprint"/>
                </StackPanel>
                <StackPanel x:Name="SecretPanel">
                  <TextBlock Text="Client secret (empty = environment variable; never written)" Style="{StaticResource MmrLabel}"/>
                  <PasswordBox x:Name="Secret"/>
                </StackPanel>
                <TextBlock x:Name="ConfigHint" Style="{StaticResource MmrHint}" Margin="0,8,0,0"/>
              </StackPanel>
            </Border>
          </Expander>
        </StackPanel>
      </ScrollViewer>

      <Grid Grid.Column="2">
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="*" MinHeight="200"/>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="200" MinHeight="110"/>
        </Grid.RowDefinitions>
        <Border Style="{StaticResource MmrCard}" Margin="0,0,0,12">
          <Grid>
            <Grid.RowDefinitions>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>
              <TextBlock Text="Mailboxes" Style="{StaticResource MmrCardTitle}"/>
              <TextBlock x:Name="Counts" Grid.Column="1" Margin="12,1,0,8" FontSize="12" VerticalAlignment="Top" Foreground="{DynamicResource TextFillColorSecondaryBrush}"/>
              <Border x:Name="StatusPill" Grid.Column="2" CornerRadius="10" Padding="10,2" VerticalAlignment="Top">
                <TextBlock x:Name="Status" FontSize="12" FontWeight="SemiBold"/>
              </Border>
            </Grid>
            <ListView x:Name="Mailboxes" Grid.Row="1" Height="150" SelectionMode="Single" BorderThickness="0" Background="Transparent" FontSize="13"
                      VirtualizingPanel.IsVirtualizing="True" VirtualizingPanel.VirtualizationMode="Recycling" ScrollViewer.HorizontalScrollBarVisibility="Disabled"
                      ItemContainerStyle="{StaticResource MmrRow}">
              <ListView.View>
                <GridView AllowsColumnReorder="False">
                  <GridViewColumn Header="Mailbox" DisplayMemberBinding="{Binding Mailbox}" Width="260"/>
                  <GridViewColumn Header="Has archive" DisplayMemberBinding="{Binding Archive}" Width="92"/>
                  <GridViewColumn Header="Folders" DisplayMemberBinding="{Binding Folders}" Width="80"/>
                  <GridViewColumn Header="Primary" DisplayMemberBinding="{Binding Primary}" Width="80"/>
                  <GridViewColumn Header="Archive" DisplayMemberBinding="{Binding ArchiveMessages}" Width="80"/>
                  <GridViewColumn Header="Recoverable" DisplayMemberBinding="{Binding Recoverable}" Width="96"/>
                  <GridViewColumn Header="Status" DisplayMemberBinding="{Binding State}" Width="120"/>
                </GridView>
              </ListView.View>
            </ListView>
          </Grid>
        </Border>
        <Border Grid.Row="1" Style="{StaticResource MmrCard}" Margin="0,0,0,6">
          <Grid>
            <Grid.RowDefinitions>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="*"/>
            </Grid.RowDefinitions>
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
              </Grid.ColumnDefinitions>
              <TextBlock Text="Preview" Style="{StaticResource MmrCardTitle}"/>
              <TextBlock x:Name="PreviewInfo" Grid.Column="1" Margin="12,1,0,8" FontSize="12" VerticalAlignment="Top" TextTrimming="CharacterEllipsis" Foreground="{DynamicResource TextFillColorSecondaryBrush}"/>
            </Grid>
            <!-- ListView/GridView rather than DataGrid: less layout per row, the list scrolls with thousands of rows.
                 Rows are compiled objects (MailboxMessageReportNative.PreviewRow). -->
            <ListView x:Name="Preview" Grid.Row="1" SelectionMode="Single" BorderThickness="0" Background="Transparent" FontSize="13"
                      VirtualizingPanel.IsVirtualizing="True" VirtualizingPanel.VirtualizationMode="Recycling" VirtualizingPanel.ScrollUnit="Item"
                      ScrollViewer.HorizontalScrollBarVisibility="Disabled" ItemContainerStyle="{StaticResource MmrRow}">
              <ListView.View>
                <GridView AllowsColumnReorder="False">
                  <GridViewColumn Header="Received" DisplayMemberBinding="{Binding Received}" Width="142"/>
                  <GridViewColumn Header="Mailbox" DisplayMemberBinding="{Binding Mailbox}" Width="0"/>
                  <GridViewColumn Header="Location" DisplayMemberBinding="{Binding Location}" Width="100"/>
                  <GridViewColumn Header="Folder" DisplayMemberBinding="{Binding FolderPath}" Width="150"/>
                  <GridViewColumn Header="Subject" DisplayMemberBinding="{Binding Subject}" Width="240"/>
                </GridView>
              </ListView.View>
            </ListView>
            <TextBlock x:Name="PreviewEmpty" Grid.Row="1" Margin="4,48,4,0" TextWrapping="Wrap" HorizontalAlignment="Center" FontSize="13" Foreground="{DynamicResource TextFillColorTertiaryBrush}"
                       Text="Type the mailboxes, the period or the subjects, then Read the messages. The window shows the first messages; the report holds them all."/>
          </Grid>
        </Border>
        <GridSplitter Grid.Row="2" Height="6" HorizontalAlignment="Stretch" Background="Transparent" ResizeBehavior="PreviousAndNext"/>
        <Border Grid.Row="3" Style="{StaticResource MmrCard}" Margin="0,6,0,12" Padding="18,10,18,10">
          <Grid>
            <Grid.RowDefinitions>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="*"/>
            </Grid.RowDefinitions>
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>
              <TextBlock Text="Progress" Style="{StaticResource MmrCardTitle}" Margin="0,0,14,6"/>
              <TextBlock x:Name="ProgressText" Grid.Column="1" Margin="0,2,12,6" FontSize="12" VerticalAlignment="Center" TextTrimming="CharacterEllipsis" Foreground="{DynamicResource TextFillColorSecondaryBrush}"/>
              <TextBlock x:Name="ProgressInfo" Grid.Column="2" Margin="0,2,0,6" FontSize="12" FontWeight="SemiBold" VerticalAlignment="Center" Foreground="{DynamicResource TextFillColorPrimaryBrush}"/>
            </Grid>
            <ProgressBar x:Name="ProgressBar" Grid.Row="1" Height="4" Minimum="0" Maximum="1" Margin="0,0,0,8" Visibility="Collapsed"/>
            <ScrollViewer x:Name="LogScroll" Grid.Row="2" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
              <ItemsControl x:Name="Log" Margin="0,0,12,0">
                <ItemsControl.ItemTemplate>
                  <DataTemplate>
                    <Grid Margin="{Binding Margin}">
                      <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="22"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                      </Grid.ColumnDefinitions>
                      <TextBlock Text="{Binding Glyph}" Foreground="{Binding Brush}" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="12" Margin="0,3,0,0" VerticalAlignment="Top"/>
                      <TextBlock Grid.Column="1" Text="{Binding Text}" TextWrapping="Wrap" FontSize="{Binding Size}" FontWeight="{Binding Weight}" Foreground="{Binding TextBrush}"/>
                      <TextBlock Grid.Column="2" Text="{Binding Time}" FontSize="11" Margin="10,2,0,0" Foreground="{DynamicResource TextFillColorTertiaryBrush}"/>
                    </Grid>
                  </DataTemplate>
                </ItemsControl.ItemTemplate>
              </ItemsControl>
            </ScrollViewer>
          </Grid>
        </Border>
      </Grid>
    </Grid>

    <Border x:Name="Actions" Grid.Row="2" Padding="24,12" BorderThickness="0,1,0,0"
            BorderBrush="{DynamicResource DividerStrokeColorDefaultBrush}" Background="{DynamicResource LayerFillColorDefaultBrush}">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <StackPanel Orientation="Horizontal">
          <Button x:Name="Search" MinWidth="190" Padding="16,6" Margin="0,0,8,0">
            <StackPanel Orientation="Horizontal">
              <TextBlock Style="{StaticResource MmrIcon}" Text="&#xE721;" Margin="0,2,8,0"/>
              <TextBlock Text="Read the messages"/>
            </StackPanel>
          </Button>
          <Button x:Name="Stop" Content="Stop" MinWidth="80" IsEnabled="False"/>
        </StackPanel>
        <TextBlock x:Name="Footer" Grid.Column="1" Margin="16,0" VerticalAlignment="Center" TextTrimming="CharacterEllipsis" FontSize="12"
                   Foreground="{DynamicResource TextFillColorSecondaryBrush}"/>
        <StackPanel Grid.Column="2" Orientation="Horizontal">
          <Button x:Name="OpenReport" Content="Open the report" Margin="0,0,8,0" IsEnabled="False"/>
          <Button x:Name="OpenCsv" Content="Open the CSV" Margin="0,0,8,0" IsEnabled="False"/>
          <Button x:Name="OpenFolder" Content="Open the folder" Margin="0,0,8,0" IsEnabled="False"/>
          <Button x:Name="Close" Content="Close" MinWidth="90" IsCancel="True"/>
        </StackPanel>
      </Grid>
    </Border>
  </Grid>
</Window>
'@
}

function New-MmrGuiBrush {
    param([Parameter(Mandatory = $true)][string]$Color)
    $brush = [Windows.Media.SolidColorBrush]::new([Windows.Media.ColorConverter]::ConvertFromString($Color))
    $brush.Freeze()
    return $brush
}

function Test-MmrGuiDarkMode {
    <# Windows shows the applications in dark mode (AppsUseLightTheme = 0); light when the setting is missing. #>
    $personalize = Get-ItemProperty -LiteralPath 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize' -ErrorAction SilentlyContinue
    return [string](Get-MmrProperty $personalize 'AppsUseLightTheme') -eq '0'
}

function Initialize-MmrGuiTheme {
    <#
        Loads WPF and applies the theme to the application: Fluent (.NET 9+), light or dark as Windows
        (System), or Light / Dark for the documentation images. Returns Fluent and Dark.
    #>
    param([ValidateSet('System', 'Light', 'Dark')][string]$Theme = 'System')

    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml
    $dark = $Theme -eq 'Dark'
    if ($Theme -eq 'System') { $dark = Test-MmrGuiDarkMode }
    # One application per process: created once, never shut down by a closed window.
    $app = [Windows.Application]::Current
    if (-not $app) {
        $app = [Windows.Application]::new()
        $app.ShutdownMode = [Windows.ShutdownMode]::OnExplicitShutdown
    }
    $fluent = $null -ne [Windows.Application].GetProperty('ThemeMode')
    if ($fluent) {
        # ThemeMode is the Fluent theme of WPF (.NET 9 and later). Light or Dark, never System: without the
        # setting (Windows Server 2016) WPF would pick dark and the colours of the window would not match.
        $app.ThemeMode = [Windows.ThemeMode]::new($(if ($dark) { 'Dark' } else { 'Light' }))
    }
    [pscustomobject]@{ Fluent = $fluent; Dark = $dark; Application = $app }
}

function Set-MmrGuiTheme {
    <#
        Colours of the window: the accent of the report on the Fluent accent resources, the status colours,
        and with the classic theme (.NET 8) the Fluent resources the window uses, with the report palette.
    #>
    param([Parameter(Mandatory = $true)][Windows.Window]$Window, [Parameter(Mandatory = $true)][pscustomobject]$Theme)

    $dark = $Theme.Dark
    $r = $Window.Resources
    $set = { param([string[]]$Keys, [string]$LightColor, [string]$DarkColor) $b = New-MmrGuiBrush $(if ($dark) { $DarkColor } else { $LightColor }); foreach ($k in $Keys) { $r[$k] = $b } }
    if (-not $Theme.Fluent) {
        & $set 'ApplicationBackgroundBrush' '#F7F4EF' '#202020'
        & $set 'CardBackgroundFillColorDefaultBrush' '#FFFFFF' '#2B2B2B'
        & $set 'CardStrokeColorDefaultBrush', 'ControlStrokeColorDefaultBrush' '#DEDEDE' '#3D3D3D'
        & $set 'ControlFillColorDefaultBrush' '#FFFFFF' '#2D2D2D'
        & $set 'ControlFillColorSecondaryBrush' '#F5F5F5' '#323232'
        & $set 'DividerStrokeColorDefaultBrush' '#DEDEDE' '#3D3D3D'
        & $set 'LayerFillColorDefaultBrush' '#FCFBF8' '#262626'
        & $set 'TextFillColorPrimaryBrush' '#242424' '#FFFFFF'
        & $set 'TextFillColorSecondaryBrush' '#5C5C5C' '#C5C5C5'
        & $set 'TextFillColorTertiaryBrush' '#8A8A8A' '#9A9A9A'
    }
    # The accent of the report instead of the accent colour of Windows.
    & $set 'AccentFillColorDefaultBrush', 'AccentButtonBackground', 'AccentButtonBorderBrush' '#B11F4B' '#FD8EA1'
    & $set 'AccentFillColorSecondaryBrush', 'AccentButtonBackgroundPointerOver' '#E6B11F4B' '#E6FD8EA1'
    & $set 'AccentFillColorTertiaryBrush', 'AccentButtonBackgroundPressed' '#CCB11F4B' '#CCFD8EA1'
    & $set 'AccentTextFillColorPrimaryBrush' '#9A1A41' '#FD8EA1'
    # Boxes, radio buttons, progress bar, calendar of the date pickers: the same accent.
    & $set 'CheckBoxCheckBackgroundFillChecked', 'CheckBoxCheckBackgroundStrokeChecked', 'RadioButtonOuterEllipseCheckedFill', 'RadioButtonOuterEllipseCheckedStroke',
        'ProgressBarForeground', 'CalendarViewSelectedBackground', 'CalendarViewSelectedBorderBrush' '#B11F4B' '#FD8EA1'
    & $set 'CheckBoxCheckBackgroundFillCheckedPointerOver', 'CheckBoxCheckBackgroundStrokeCheckedPointerOver', 'RadioButtonOuterEllipseCheckedStrokePointerOver' '#E6B11F4B' '#E6FD8EA1'
    & $set 'CheckBoxCheckBackgroundFillCheckedPressed', 'CheckBoxCheckBackgroundStrokeCheckedPressed' '#CCB11F4B' '#CCFD8EA1'
    # Row selected in the lists: a soft accent, the text stays readable.
    & $set 'DataGridRowSelectedBackgroundThemeBrush' '#1FB11F4B' '#40FD8EA1'
    # The mark of the row selected in the lists (Fluent): the accent of the report.
    & $set 'ListViewItemPillFillBrush' '#B11F4B' '#FD8EA1'
    & $set 'DataGridRowSelectedForegroundThemeBrush' '#242424' '#FFFFFF'
    $r[[Windows.SystemColors]::HighlightBrushKey] = $r['DataGridRowSelectedBackgroundThemeBrush']
    $r[[Windows.SystemColors]::InactiveSelectionHighlightBrushKey] = $r['DataGridRowSelectedBackgroundThemeBrush']
    $r[[Windows.SystemColors]::HighlightTextBrushKey] = $r['DataGridRowSelectedForegroundThemeBrush']
    $r[[Windows.SystemColors]::InactiveSelectionHighlightTextBrushKey] = $r['DataGridRowSelectedForegroundThemeBrush']
    & $set 'MmrBrand' '#B11F4B' '#B11F4B'
    & $set 'MmrBrandText' '#B11F4B' '#FD8EA1'
    & $set 'MmrAccentSoft' '#14B11F4B' '#33FD8EA1'
    & $set 'MmrSuccess' '#16A34A' '#4ADE80'
    & $set 'MmrCaution' '#D97706' '#FBBF24'
    & $set 'MmrCritical' '#DC2626' '#F87171'
    & $set 'MmrInfoBackground' '#F3F3F3' '#2E2E2E'
    & $set 'MmrCautionBackground' '#FFF7E8' '#33FBBF24'
    & $set 'MmrCriticalBackground' '#FDECEC' '#33F87171'
    & $set 'MmrSuccessBackground' '#EAF7EE' '#334ADE80'
}

function Invoke-MmrGuiPump {
    <# Lets the window repaint and handle clicks during a run (the WPF equivalent of DoEvents). #>
    $frame = [Windows.Threading.DispatcherFrame]::new()
    [void][Windows.Threading.Dispatcher]::CurrentDispatcher.BeginInvoke([Windows.Threading.DispatcherPriority]::Background,
        [Windows.Threading.DispatcherOperationCallback] { param($f) $f.Continue = $false; $null }, $frame)
    [Windows.Threading.Dispatcher]::PushFrame($frame)
}

function New-MmrForm {
    <#
    .SYNOPSIS
        Builds the window (without showing it). Used by Show-MmrGui, the tests and the documentation tool.
    .PARAMETER Theme
        System (like Windows), Light or Dark (documentation images, tests).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][hashtable]$Configuration, [ValidateSet('System', 'Light', 'Dark')][string]$Theme = 'System')

    $look = Initialize-MmrGuiTheme -Theme $Theme
    $window = [Windows.Markup.XamlReader]::Parse((Get-MmrGuiXaml))
    $window.Title = "Mailbox Message Report $($script:ToolVersion)"
    Set-MmrGuiTheme -Window $window -Theme $look
    $controls = @{}
    foreach ($name in 'Root', 'Header', 'Version', 'SettingsScroll', 'Inputs', 'Mailbox', 'LoadMailboxes', 'MailboxHint', 'StartDate', 'EndDate', 'Subject',
        'Recipients', 'LocationPrimary', 'LocationArchive', 'RecoverableItems', 'ExcludeFolders', 'FormatCsv', 'FormatHtml', 'LayoutGlobal', 'LayoutPerMailbox', 'LayoutBoth', 'ReportHint',
        'ConnectionExpander', 'TenantId', 'AppId', 'AuthMode', 'ThumbPanel', 'Thumbprint', 'SecretPanel', 'Secret', 'ConfigHint', 'Counts', 'StatusPill', 'Status',
        'Mailboxes', 'Preview', 'PreviewInfo', 'PreviewEmpty', 'ProgressBar', 'ProgressText', 'ProgressInfo', 'LogScroll', 'Log', 'Actions', 'Search', 'Stop', 'Footer',
        'OpenReport', 'OpenCsv', 'OpenFolder', 'Close') {
        $controls[$name] = $window.FindName($name)
    }
    if ($look.Fluent) { $controls.Search.SetResourceReference([Windows.FrameworkElement]::StyleProperty, 'AccentButtonStyle') }
    else { $controls.Search.SetResourceReference([Windows.Controls.Control]::BackgroundProperty, 'AccentFillColorDefaultBrush'); $controls.Search.Foreground = [Windows.Media.Brushes]::White }
    $controls.Version.Text = "v$($script:ToolVersion)  " + [char]0x00B7 + '  Nicolas Fabert'
    # The button of the window in the taskbar shows the progress of a run too.
    $window.TaskbarItemInfo = [Windows.Shell.TaskbarItemInfo]::new()

    # Values of the configuration.
    if ([int]$Configuration.PastDays -gt 0) {
        $zone = Get-MmrTimeZone $Configuration.TimeZone
        $controls.StartDate.SelectedDate = [TimeZoneInfo]::ConvertTimeFromUtc([datetime]::UtcNow, $zone).Date.AddDays(-[int]$Configuration.PastDays)
    }
    $controls.LocationPrimary.IsChecked = @($Configuration.Locations) -contains 'Primary'
    $controls.LocationArchive.IsChecked = @($Configuration.Locations) -contains 'Archive'
    $controls.RecoverableItems.IsChecked = [bool]$Configuration.RecoverableItems
    $controls.Recipients.IsChecked = [bool]$Configuration.Recipients
    $controls.ExcludeFolders.Text = (@($Configuration.ExcludeFolders) -join [Environment]::NewLine)
    $controls.FormatCsv.IsChecked = @($Configuration.ReportFormats) -contains 'Csv'
    $controls.FormatHtml.IsChecked = @($Configuration.ReportFormats) -contains 'Html'
    switch ([string]$Configuration.ReportLayout) { 'PerMailbox' { $controls.LayoutPerMailbox.IsChecked = $true } 'Both' { $controls.LayoutBoth.IsChecked = $true } default { $controls.LayoutGlobal.IsChecked = $true } }
    $controls.TenantId.Text = [string]$Configuration.TenantId
    $controls.AppId.Text = [string]$Configuration.AppId
    foreach ($m in 'Certificate', 'ClientSecret') { [void]$controls.AuthMode.Items.Add($m) }
    $controls.AuthMode.SelectedItem = [string]$Configuration.AuthMode
    $controls.Thumbprint.Text = [string]$Configuration.CertificateThumbprint
    $controls.ConfigHint.Text = "From $([string](Get-MmrProperty $Configuration 'ConfigPath')). Changes here are for this window only."
    $controls.ConnectionExpander.IsExpanded = -not ($Configuration.TenantId -and $Configuration.AppId)

    # Lists replaced in one go (one refresh), rows compiled (MailboxMessageReportNative.MailboxRow / PreviewRow).
    $mailboxRows = [MailboxMessageReportNative.BulkCollection]::new()
    $previewRows = [MailboxMessageReportNative.BulkCollection]::new()
    $items = [Collections.ObjectModel.ObservableCollection[object]]::new()
    $controls.Mailboxes.ItemsSource = $mailboxRows
    $controls.Preview.ItemsSource = $previewRows
    $controls.Log.ItemsSource = $items
    # The channel of the background run (Start-MmrGuiWork): its lines, Stop (Cancel), the log file. Every key the engine
    # reads is there (a synchronized hashtable throws on a missing key under Set-StrictMode).
    $shared = [hashtable]::Synchronized(@{ Cancel = $false; Queue = [Collections.Concurrent.ConcurrentQueue[string[]]]::new(); Log = $null; Sink = $null; Pump = $null })
    $timer = [Windows.Threading.DispatcherTimer]::new([Windows.Threading.DispatcherPriority]::Background)
    $timer.Interval = [TimeSpan]::FromMilliseconds(100)
    $timer.Add_Tick({ Step-MmrGuiWork })
    $script:Gui = @{
        Form = $window; Controls = $controls; Configuration = $Configuration.Clone(); Settings = $null; Theme = $look
        Running = $false; Result = $null; LastReport = $null; LastCsv = $null; LastFolder = $null
        ArchiveGuids = @{}; MailboxRows = $mailboxRows; PreviewRows = $previewRows; Items = $items; Lines = [Collections.Generic.List[string]]::new()
        Shared = $shared; Timer = $timer; Job = $null; Runspace = $null
        # The progress of the run in course (Set-MmrGuiProgress): its step, its start, the part done (-1: none yet).
        Progress = @{ Active = $false; Step = ''; Started = [datetime]::UtcNow; Fraction = -1.0; Stopping = $false }
        # Attached to Closing only while a run is in progress: closing then stops the run first.
        ClosingGuard = [ComponentModel.CancelEventHandler] {
            param($sender, $e)
            $e.Cancel = $true
            if ($script:Gui) { $script:Gui.Shared.Cancel = $true; Set-MmrGuiProgress -Stopping }
            Add-MmrGuiLine 'Warn' 'A run is in progress: it stops at the next page, then the window can be closed.'
        }
    }
    Set-MmrGuiStatus 'Ready' 'Ready'
    $controls.Footer.Text = "Reports: $($Configuration.OutputPath)"

    $controls.Search.Add_Click({ Invoke-MmrGuiSearch })
    $controls.Stop.Add_Click({
            $g = $script:Gui
            if ($g -and $g.Running) {
                $g.Shared.Cancel = $true
                Set-MmrGuiProgress -Stopping
                Add-MmrGuiLine 'Warn' 'Stop requested: the run stops at the next page.'
            }
        })
    # A column header: the list is sorted. GridView has no proportional width: the subject takes the width left.
    foreach ($list in $controls.Mailboxes, $controls.Preview) {
        $list.AddHandler([Windows.Controls.Primitives.ButtonBase]::ClickEvent, [Windows.RoutedEventHandler] { param($sender, $e) if ($e.OriginalSource -is [Windows.Controls.GridViewColumnHeader]) { Set-MmrGuiSort $sender $e.OriginalSource } })
        $list.Add_SizeChanged({ Update-MmrGuiColumns })
    }
    $controls.AuthMode.Add_SelectionChanged({ Update-MmrGuiState })
    foreach ($c in 'LocationPrimary', 'LocationArchive', 'FormatCsv', 'FormatHtml', 'LayoutGlobal', 'LayoutPerMailbox', 'LayoutBoth') { $controls[$c].Add_Click({ Update-MmrGuiState }) }
    $controls.LoadMailboxes.Add_Click({ Import-MmrGuiMailboxes })
    $controls.Mailbox.Add_TextChanged({ Update-MmrGuiMailboxHint })
    $controls.OpenReport.Add_Click({ if ($script:Gui.LastReport) { Start-Process -FilePath $script:Gui.LastReport } })
    $controls.OpenCsv.Add_Click({ if ($script:Gui.LastCsv) { Start-Process -FilePath $script:Gui.LastCsv } })
    $controls.OpenFolder.Add_Click({ if ($script:Gui.LastFolder) { Start-Process -FilePath $script:Gui.LastFolder } })

    # The window fits the screen where it opens (small laptop screen at 150 %); the left column scrolls.
    $area = [Windows.SystemParameters]::WorkArea
    $window.Width = [Math]::Min($window.Width, $area.Width)
    $window.Height = [Math]::Min($window.Height, $area.Height)
    $window.MinWidth = [Math]::Min($window.MinWidth, $area.Width)
    $window.MinHeight = [Math]::Min($window.MinHeight, $area.Height)
    Update-MmrGuiMailboxHint
    Update-MmrGuiState
    [pscustomobject]@{ Form = $window; Controls = $controls; Lines = $script:Gui.Lines; Items = $items; Mailboxes = $mailboxRows; Preview = $previewRows }
}

function Update-MmrGuiState {
    <# What can be used now: the read button, the secret, the hint of the report. #>
    $g = $script:Gui
    if (-not $g) { return }
    $c = $g.Controls
    $secret = [string]$c.AuthMode.SelectedItem -eq 'ClientSecret'
    $c.SecretPanel.Visibility = if ($secret) { 'Visible' } else { 'Collapsed' }
    $c.ThumbPanel.Visibility = if ($secret) { 'Collapsed' } else { 'Visible' }
    $where = [bool]$c.LocationPrimary.IsChecked -or [bool]$c.LocationArchive.IsChecked
    $formats = [bool]$c.FormatCsv.IsChecked -or [bool]$c.FormatHtml.IsChecked
    $c.Search.IsEnabled = -not $g.Running -and $where -and $formats
    $c.ReportHint.Text = if (-not $formats) { 'Tick CSV, HTML or both.' }
    elseif ($c.LayoutPerMailbox.IsChecked) { 'A folder Mailboxes\ with the CSV and HTML files of each mailbox, and a summary with a link to each one.' }
    elseif ($c.LayoutBoth.IsChecked) { 'One report for every mailbox, and the files of each mailbox in Mailboxes\.' }
    else { "One file for every mailbox. A HTML report shows the first $('{0:N0}' -f [int]$g.Configuration.HtmlMaxMessages) messages; the CSV file holds them all." }
}

function Update-MmrGuiMailboxHint {
    <# Under the mailboxes: how many addresses are typed, which are not valid, how many have an ArchiveGuid of a list. #>
    $g = $script:Gui
    if (-not $g) { return }
    $list = @(Split-MmrList @($g.Controls.Mailbox.Text) -Spaces)
    if (-not $list.Count) { $g.Controls.MailboxHint.Text = 'One per line (or separated by ;): SMTP address (any alias) or UPN. A list: text or CSV file (PrimarySmtpAddress, and ArchiveGuid to read the archive without User.Read.All).'; return }
    $bad = @($list | Where-Object { $_ -notmatch $script:SmtpPattern })
    $guids = @($list | Where-Object { $g.ArchiveGuids.ContainsKey($_.ToLowerInvariant()) }).Count
    $g.Controls.MailboxHint.Text = '{0} mailbox{1}{2}{3}' -f $list.Count, $(if ($list.Count -gt 1) { 'es' } else { '' }), $(if ($guids) { " $($script:Dot) $guids with the ArchiveGuid of the list" } else { '' }), $(if ($bad.Count) { " $($script:Dot) not an address: $(($bad | Select-Object -First 3) -join ', ')" } else { '' })
}

function Import-MmrGuiMailboxes {
    <# Load a list of mailboxes (text or CSV file) into the box; the ArchiveGuid of a CSV file is kept for the run. #>
    $g = $script:Gui
    $dialog = [Microsoft.Win32.OpenFileDialog]::new()
    $dialog.Filter = 'Mailboxes (*.txt;*.csv)|*.txt;*.csv|All files (*.*)|*.*'
    $dialog.Title = 'List of mailboxes: one address per line, or a CSV file (PrimarySmtpAddress, UserPrincipalName, Mail... and ArchiveGuid)'
    if (-not $dialog.ShowDialog($g.Form)) { return }
    try {
        $list = @(Read-MmrMailboxFile -Path $dialog.FileName)
        if (-not $list.Count) { Add-MmrGuiLine 'Warn' "No address in $($dialog.FileName)."; return }
        foreach ($e in $list) { if ($e.ArchiveGuid) { $g.ArchiveGuids[$e.Address] = $e.ArchiveGuid } }
        $g.Controls.Mailbox.Text = (@($list | ForEach-Object Address) -join [Environment]::NewLine)
        $withGuid = @($list | Where-Object ArchiveGuid).Count
        Add-MmrGuiLine 'Info' "$($list.Count) mailbox(es) loaded from $($dialog.FileName)$(if ($withGuid) { ", $withGuid with an ArchiveGuid" })."
    }
    catch { Add-MmrGuiLine 'Fail' $_.Exception.Message }
}

function Set-MmrGuiSort {
    <# A column header clicked: the list sorted by that column, ascending then descending. #>
    param($List, $Header)
    $column = $Header.Column
    if (-not $column -or -not $column.DisplayMemberBinding) { return }
    $path = $column.DisplayMemberBinding.Path.Path
    $view = [Windows.Data.CollectionViewSource]::GetDefaultView($List.ItemsSource)
    $direction = [ComponentModel.ListSortDirection]::Ascending
    if ($view.SortDescriptions.Count -and $view.SortDescriptions[0].PropertyName -eq $path -and $view.SortDescriptions[0].Direction -eq $direction) { $direction = [ComponentModel.ListSortDirection]::Descending }
    $view.SortDescriptions.Clear()
    $view.SortDescriptions.Add([ComponentModel.SortDescription]::new($path, $direction))
}

function Update-MmrGuiColumns {
    <#
        GridView has no proportional width: in the preview the folder and the subject share the width left (35 / 65 %),
        in the list of mailboxes the mailbox takes it.
    #>
    $g = $script:Gui
    if (-not $g) { return }
    foreach ($pair in @(@($g.Controls.Preview, @{ Folder = 0.35; Subject = 0.65 }), @($g.Controls.Mailboxes, @{ Mailbox = 1.0 }))) {
        $list = $pair[0]; $share = $pair[1]
        if ($list.ActualWidth -le 0) { continue }
        $fixed = 0.0
        foreach ($col in $list.View.Columns) { if (-not $share.ContainsKey([string]$col.Header) -and -not [double]::IsNaN($col.Width)) { $fixed += $col.Width } }
        $free = [Math]::Max(200, $list.ActualWidth - $fixed - 34)
        foreach ($col in $list.View.Columns) { if ($share.ContainsKey([string]$col.Header)) { $col.Width = [Math]::Max(90, [Math]::Floor($free * $share[[string]$col.Header])) } }
    }
}

function Add-MmrGuiLine {
    <# One line of the progress: icon and colour of its status. 'Progress' updates the progress bar instead. #>
    param([string]$Status, [string]$Text, [switch]$NoScroll)

    $g = $script:Gui
    if (-not $g) { return }
    if ($Status -eq 'Progress') {
        # fraction|text|time left (Write-MmrProgress), cut at the first and the last '|'; the time left may be absent.
        $fraction = 0.0; $label = $Text; $left = ''
        $first = $Text.IndexOf('|'); $last = $Text.LastIndexOf('|')
        if ($first -gt 0 -and [double]::TryParse($Text.Substring(0, $first), [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$fraction)) {
            if ($last -gt $first) { $label = $Text.Substring($first + 1, $last - $first - 1); $left = $Text.Substring($last + 1) }
            else { $label = $Text.Substring($first + 1) }
        }
        Set-MmrGuiProgress -Fraction $fraction -Text $label -Left $left
        return
    }
    $glyphs = @{ Step = 0xE76C; Ok = 0xE73E; Warn = 0xE7BA; Fail = 0xEA39; Info = 0xE946; Skip = 0xE72A }
    $colours = @{ Step = 'MmrBrandText'; Ok = 'MmrSuccess'; Warn = 'MmrCaution'; Fail = 'MmrCritical'; Info = 'TextFillColorSecondaryBrush'; Skip = 'TextFillColorTertiaryBrush' }
    $key = if ($glyphs.ContainsKey($Status)) { $Status } else { 'Info' }
    $step = $Status -eq 'Step'
    if ($step) { Set-MmrGuiProgress -Step $Text }
    $shown = if ($step) { $Text -replace '^\[(\d+/\d+)\]\s*', '$1   ' } else { $Text }
    $window = $g.Form
    $g.Items.Add([pscustomobject]@{
            Glyph     = [string][char]$glyphs[$key]
            Brush     = $window.TryFindResource($colours[$key])
            Text      = $shown
            TextBrush = $window.TryFindResource($(if ($key -in 'Info', 'Skip') { 'TextFillColorSecondaryBrush' } else { 'TextFillColorPrimaryBrush' }))
            Weight    = if ($step) { [Windows.FontWeights]::SemiBold } else { [Windows.FontWeights]::Normal }
            Size      = if ($step) { 13 } else { 12 }
            Margin    = if ($step) { [Windows.Thickness]::new(0, $(if ($g.Items.Count) { 10 } else { 0 }), 0, 3) } else { [Windows.Thickness]::new(0, 1, 0, 1) }
            Time      = (Get-Date).ToString('HH:mm:ss')
        })
    $g.Lines.Add("[$Status] $Text")
    if (-not $NoScroll) { $g.Controls.LogScroll.ScrollToEnd() }
}

function Set-MmrGuiProgress {
    <#
        The progress bar of the window and its button in the taskbar, during a run:
          -Start <text>  the run begins: the bar moves (nothing counted yet), the time since the start on the right;
          -Step <text>   a step begins ('[3/6] Title'): the same, with the step;
          -Fraction      the part done, its text (1,240/1,858 folders read) and the time left (Write-MmrProgress);
          -Tick          the timer of the run (Step-MmrGuiWork): the time since the start while nothing is counted;
          -Stopping      Stop requested: the taskbar button turns yellow, 'Stopping...';
          -Waiting       a question to the administrator: the bar stops;
          -Done          the run is over: bar and texts hidden, taskbar button back to normal.
        The window reads the lines of a run every 100 ms and keeps the last part done only: ten updates a second at
        most, whatever the volume.
    #>
    [CmdletBinding(DefaultParameterSetName = 'Tick')]
    param(
        [Parameter(ParameterSetName = 'Start', Mandatory = $true)][string]$Start,
        [Parameter(ParameterSetName = 'Step', Mandatory = $true)][string]$Step,
        [Parameter(ParameterSetName = 'Fraction', Mandatory = $true)][double]$Fraction,
        [Parameter(ParameterSetName = 'Fraction')][AllowEmptyString()][string]$Text,
        [Parameter(ParameterSetName = 'Fraction')][AllowEmptyString()][string]$Left,
        [Parameter(ParameterSetName = 'Tick')][switch]$Tick,
        [Parameter(ParameterSetName = 'Stopping', Mandatory = $true)][switch]$Stopping,
        [Parameter(ParameterSetName = 'Waiting', Mandatory = $true)][switch]$Waiting,
        [Parameter(ParameterSetName = 'Done', Mandatory = $true)][switch]$Done
    )

    $g = $script:Gui
    if (-not $g) { return }
    $c = $g.Controls; $p = $g.Progress; $task = $g.Form.TaskbarItemInfo
    $dot = [char]0x00B7
    switch ($PSCmdlet.ParameterSetName) {
        'Start' {
            $g.Progress = $p = @{ Active = $true; Step = $Start; Started = [datetime]::UtcNow; Fraction = -1.0; Stopping = $false }
            $c.ProgressBar.Visibility = 'Visible'
        }
        'Done' {
            $p.Active = $false
            $c.ProgressBar.IsIndeterminate = $false
            $c.ProgressBar.Visibility = 'Collapsed'
            $c.ProgressText.Text = ''; $c.ProgressInfo.Text = ''
            if ($task) { $task.ProgressState = 'None' }
            return
        }
    }
    if (-not $p.Active) { return }
    switch ($PSCmdlet.ParameterSetName) {
        'Step' { $p.Step = 'Step ' + ($Step -replace '^\[(\d+/\d+)\]\s*', ('$1 ' + $dot + ' ')); $p.Fraction = -1.0; $c.ProgressBar.Visibility = 'Visible' }
        'Stopping' { $p.Stopping = $true }
        'Waiting' {
            # -2: nothing moves until the next -Start or -Step.
            $p.Fraction = -2.0
            $c.ProgressBar.IsIndeterminate = $false
            $c.ProgressBar.Visibility = 'Collapsed'
            $c.ProgressText.Text = 'Waiting for your answer'; $c.ProgressInfo.Text = ''
            if ($task) { $task.ProgressState = 'None' }
            return
        }
        'Fraction' {
            $p.Fraction = [Math]::Min(1.0, [Math]::Max(0.0, $Fraction))
            $c.ProgressBar.IsIndeterminate = $false
            $c.ProgressBar.Value = $p.Fraction
            $c.ProgressText.Text = if ($Text) { "$($p.Step)  $dot  $Text" } else { $p.Step }
            $percent = [string]::Format([Globalization.CultureInfo]::InvariantCulture, '{0:0} %', [Math]::Floor($p.Fraction * 100))
            $c.ProgressInfo.Text = if ($p.Stopping) { 'Stopping...' } elseif ($Left) { "$percent  $dot  $Left" } else { $percent }
            if ($task) { $task.ProgressState = $(if ($p.Stopping) { 'Paused' } else { 'Normal' }); $task.ProgressValue = $p.Fraction }
            return
        }
    }
    if ($p.Fraction -ge 0 -or $p.Fraction -le -2) {
        # A part is known (the bar stays where it is), or a question is asked; Stop turns the taskbar button yellow.
        if ($p.Stopping -and $p.Fraction -ge 0) { $c.ProgressInfo.Text = 'Stopping...'; if ($task) { $task.ProgressState = 'Paused' } }
        return
    }
    # Nothing counted yet in this step: the bar moves, the time since the start of the run.
    if (-not $c.ProgressBar.IsIndeterminate) { $c.ProgressBar.IsIndeterminate = $true }
    if ($c.ProgressText.Text -ne $p.Step) { $c.ProgressText.Text = $p.Step }
    $info = if ($p.Stopping) { 'Stopping...' } else {
        $t = [datetime]::UtcNow - $p.Started
        if ($t.TotalHours -ge 1) { '{0}:{1:00}:{2:00} elapsed' -f [int][Math]::Floor($t.TotalHours), $t.Minutes, $t.Seconds } else { '{0}:{1:00} elapsed' -f $t.Minutes, $t.Seconds }
    }
    if ($c.ProgressInfo.Text -ne $info) { $c.ProgressInfo.Text = $info }
    if ($task) {
        $state = if ($p.Stopping) { 'Paused' } else { 'Indeterminate' }
        if ([string]$task.ProgressState -ne $state) { $task.ProgressState = $state; if ($p.Stopping) { $task.ProgressValue = 1 } }
    }
}

function Set-MmrGuiStatus {
    <# The status pill of the run: Ready, Running, Completed, Warning or Failed. #>
    param([string]$Text, [string]$Status)

    $c = $script:Gui.Controls
    $c.Status.Text = $Text
    $pair = switch ($Status) {
        'Completed' { 'MmrSuccess', 'MmrSuccessBackground' }
        'Failed' { 'MmrCritical', 'MmrCriticalBackground' }
        'Warning' { 'MmrCaution', 'MmrCautionBackground' }
        'Running' { 'MmrBrandText', 'MmrAccentSoft' }
        default { 'TextFillColorSecondaryBrush', 'MmrInfoBackground' }
    }
    $c.Status.SetResourceReference([Windows.Controls.TextBlock]::ForegroundProperty, $pair[0])
    $c.StatusPill.SetResourceReference([Windows.Controls.Border]::BackgroundProperty, $pair[1])
}


function Get-MmrGuiSettings {
    <# The configuration with the connection fields of the window. #>
    $g = $script:Gui
    $c = $g.Controls
    $cfg = $g.Configuration.Clone()
    $cfg.TenantId = $c.TenantId.Text.Trim()
    $cfg.AppId = $c.AppId.Text.Trim()
    $cfg.AuthMode = [string]$c.AuthMode.SelectedItem
    $cfg.CertificateThumbprint = ($c.Thumbprint.Text -replace '\s', '').Trim()
    return $cfg
}

function Update-MmrGuiRows {
    <# The mailboxes and the preview from the result of the run, replaced in one go. #>
    param($Preview)
    $g = $script:Gui
    $rows = foreach ($m in @(if ($g.Result) { $g.Result.Mailboxes })) {
        $row = [MailboxMessageReportNative.MailboxRow]::new()
        $row.Mailbox = if ($m.DisplayName) { "$($m.DisplayName) <$($m.Address)>" } else { $m.Address }
        $row.Name = $m.DisplayName
        $row.Archive = switch ($m.Archive) { 'Yes' { if ($m.ArchiveSource -eq 'File') { 'Yes (list)' } else { 'Yes' } } 'No' { 'No' } default { '?' } }
        $row.Folders = if ($m.FoldersFailed) { '{0} ({1} err.)' -f $m.FoldersRead, $m.FoldersFailed } else { [string]$m.FoldersRead }
        $row.Primary = $m.PrimaryMessages; $row.ArchiveMessages = $m.ArchiveMessages; $row.Recoverable = $m.RecoverableMessages
        $row.State = if ($m.State -ne 'Ok') { "Not read: $($m.Detail)" } elseif ($m.Status -eq 'Partial') { "Partial: $($m.Notes)" } else { $m.Status }
        $row
    }
    foreach ($list in $g.Controls.Mailboxes, $g.Controls.Preview) { [Windows.Data.CollectionViewSource]::GetDefaultView($list.ItemsSource).SortDescriptions.Clear() }
    $g.MailboxRows.ReplaceAll(@($rows))
    $g.PreviewRows.ReplaceAll([MailboxMessageReportNative.PreviewRow]::Build($Preview))
    $many = @($rows).Count -gt 1
    foreach ($column in $g.Controls.Preview.View.Columns) { if ($column.Header -eq 'Mailbox') { $column.Width = if ($many) { 180 } else { 0 } } }
    Update-MmrGuiColumns
    $total = if ($g.Result) { [long]$g.Result.Counts.Messages } else { 0 }
    $shown = $g.PreviewRows.Count
    $g.Controls.PreviewEmpty.Visibility = if ($shown) { 'Collapsed' } else { 'Visible' }
    if (-not $shown -and $g.Result) { $g.Controls.PreviewEmpty.Text = 'No message: widen the period, check the subjects, or tick the archive.' }
    $g.Controls.PreviewInfo.Text = if (-not $g.Result) { '' } elseif ($total -gt $shown) { "the first $('{0:N0}' -f $shown) of $('{0:N0}' -f $total) messages $($script:Dot) every message is in the report" } else { "$('{0:N0}' -f $total) message(s)" }
    if ($g.Result) { $n = $g.Result.Counts; $g.Controls.Counts.Text = '{0:N0} read {1} {2:N0} with an archive {1} {3:N0} not read' -f ($n.MailboxesRead + $n.MailboxesPartial), $script:Dot, $n.WithArchive, $n.MailboxesNotRead }
    else { $g.Controls.Counts.Text = '' }
}

function Start-MmrGuiRun {
    <# A run starts: inputs and buttons disabled, Stop enabled, a closing of the window stops the run first. #>
    param([string]$Text)
    $g = $script:Gui
    $c = $g.Controls
    $g.Shared.Cancel = $false
    $g.Shared.Log = $script:LogWriter
    $g.Running = $true
    $g.Form.add_Closing($g.ClosingGuard)
    foreach ($b in 'Search', 'OpenReport', 'OpenCsv', 'OpenFolder', 'Close', 'Inputs') { $c[$b].IsEnabled = $false }
    $c.Stop.IsEnabled = $true
    Set-MmrGuiStatus $Text 'Running'
    Set-MmrGuiProgress -Start $Text
}

function Stop-MmrGuiRun {
    $g = $script:Gui
    $c = $g.Controls
    $g.Form.remove_Closing($g.ClosingGuard)
    $g.Running = $false
    foreach ($b in 'Close', 'Inputs') { $c[$b].IsEnabled = $true }
    $c.Stop.IsEnabled = $false
    Set-MmrGuiProgress -Done
    $c.OpenReport.IsEnabled = [bool]$g.LastReport
    $c.OpenCsv.IsEnabled = [bool]$g.LastCsv
    $c.OpenFolder.IsEnabled = [bool]$g.LastFolder
    Update-MmrGuiState
}

#region Background work ---------------------------------------------------------------------------------------
# The run of the window goes on in a runspace of its own, with the module loaded there: the window keeps answering
# whatever the volume. Its lines go through a queue (Shared.Queue) that the window reads every 100 ms; Stop goes
# through Shared too.

$script:GuiInline = $false
$script:GuiWorkScript = @'
param($Kind, $Arguments, $Shared)
& (Get-Module MailboxMessageReport) { param($Kind, $Arguments, $Shared) Invoke-MmrGuiWork -Kind $Kind -Arguments $Arguments -Shared $Shared } $Kind $Arguments $Shared
'@

function Invoke-MmrGuiWork {
    <#
    .SYNOPSIS
        The work of the window, in its background runspace (or inline): Search = connection, mailboxes, folders,
        messages, report. Never throws: returns @{ Ok; Cancelled; Error; Result; Report; Preview }.
    #>
    param([Parameter(Mandatory = $true)][string]$Kind, [hashtable]$Arguments = @{}, [Parameter(Mandatory = $true)][hashtable]$Shared)
    $script:Ui = $Shared
    $script:Quiet = $true
    if ($Shared.Log -and -not [object]::ReferenceEquals($script:LogWriter, $Shared.Log)) { $script:LogWriter = $Shared.Log }
    $a = $Arguments
    $out = @{ Ok = $false; Cancelled = $false; Error = ''; Result = $null; Report = $null; Preview = $null }
    $partsPath = $null
    try {
        if ($Kind -ne 'Search') { throw "Unknown work of the window: $Kind" }
        $s = $a.Settings
        Initialize-MmrSteps -Total 5
        Write-MmrNextStep 'Microsoft Graph' 'Key'
        $connection = Connect-MmrGraph -Settings $s -Secret $a.Secret
        Write-MmrItem Ok ('Application {0} {1} tenant {2}' -f $(if ($connection.AppName) { $connection.AppName } else { $s.AppId }), [char]0x00B7, $connection.TenantGuid) -Icon Key
        Write-MmrItem Info ('Permissions: {0}' -f (@($connection.Roles) -join ', ')) -Icon Shield
        $runPath = New-MmrRunFolder -OutputPath $s.OutputPath -Prefix $s.ReportPrefix
        $partsPath = Join-Path $runPath '.parts'
        $out.Result = Find-MmrMessages -Settings $s -Request $a.Request -PartsPath $partsPath
        Write-MmrNextStep 'Report' 'Report'
        $report = Export-MmrReport -Result $out.Result -Directory $runPath -Prefix $s.ReportPrefix -Formats $a.Request.Formats -Layout $a.Request.Layout -Delimiter $s.CsvDelimiter -HtmlMaxMessages $s.HtmlMaxMessages -PreviewMessages $s.PreviewMessages -PartsPath $partsPath
        Write-MmrItem Ok "Report: $($report.Directory)" -Icon File
        $csv = if ($report.Files.Contains('Messages')) { $report.Files.Messages } elseif ($report.Files.Contains('Mailboxes')) { $report.Files.Mailboxes } else { $null }
        $out.Report = @{ Directory = $report.Directory; Html = Get-MmrProperty $report.Files 'Html'; Csv = $csv }
        $out.Preview = $report.Preview
        $out.Ok = $true
    }
    catch [OperationCanceledException] { $out.Cancelled = $true }
    catch {
        $out.Error = $_.Exception.Message
        Write-MmrLog 'ERROR' "Window ($Kind): $($_.Exception.Message)"
    }
    finally {
        if (-not $out.Ok -and $partsPath -and (Test-Path -LiteralPath $partsPath)) { Remove-Item -LiteralPath $partsPath -Recurse -Force -ErrorAction SilentlyContinue }
        $script:Ui = $null
    }
    return $out
}

function Open-MmrGuiRunspace {
    <# The background runspace of the window, opened in the background (module loaded there) when the window opens. #>
    $g = $script:Gui
    if ($g.Runspace) { return }
    $iss = [Management.Automation.Runspaces.InitialSessionState]::CreateDefault2()
    $iss.ImportPSModule([string[]]@(Join-Path $script:ToolRoot 'MailboxMessageReport.psd1'))
    $runspace = [Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace($iss)
    # STA, like the thread of the window.
    $runspace.ApartmentState = [Threading.ApartmentState]::STA
    $runspace.ThreadOptions = [Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
    $runspace.OpenAsync()
    $g.Runspace = $runspace
}

function Get-MmrGuiRunspace {
    <# The background runspace, once open and free (the module is loaded while it is still Busy after Opened). #>
    $g = $script:Gui
    if (-not $g.Runspace) { Open-MmrGuiRunspace }
    $until = [datetime]::UtcNow.AddSeconds(90)
    while (($g.Runspace.RunspaceStateInfo.State -in 'BeforeOpen', 'Opening' -or ($g.Runspace.RunspaceStateInfo.State -eq 'Opened' -and $g.Runspace.RunspaceAvailability -ne 'Available')) -and [datetime]::UtcNow -lt $until) {
        Invoke-MmrGuiPump; Start-Sleep -Milliseconds 30
    }
    if ($g.Runspace.RunspaceStateInfo.State -ne 'Opened') { throw "The engine of the window could not start: $($g.Runspace.RunspaceStateInfo.Reason)" }
    if ($g.Runspace.RunspaceAvailability -ne 'Available') { throw 'The engine of the window is still busy: try again in a moment.' }
    return $g.Runspace
}

function Start-MmrGuiWork {
    <#
        Runs one piece of work (Invoke-MmrGuiWork) in the background runspace of the window and returns at once;
        OnDone runs on the window thread with the outcome and the Context. Inline ($script:GuiInline: the tests and
        the documentation tool, whose simulated tenant lives in this runspace): the same work, on this thread.
    #>
    param([Parameter(Mandatory = $true)][string]$Kind, [hashtable]$Arguments = @{}, [Parameter(Mandatory = $true)][scriptblock]$OnDone, [hashtable]$Context = @{})
    $g = $script:Gui
    if ($script:GuiInline) {
        $outcome = Invoke-MmrGuiWork -Kind $Kind -Arguments $Arguments -Shared $g.Shared
        Receive-MmrGuiMessages
        & $OnDone $outcome $Context
        return
    }
    $ps = [PowerShell]::Create()
    $ps.Runspace = Get-MmrGuiRunspace
    [void]$ps.AddScript($script:GuiWorkScript).AddArgument($Kind).AddArgument($Arguments).AddArgument($g.Shared)
    $g.Job = @{ PowerShell = $ps; Handle = $ps.BeginInvoke(); OnDone = $OnDone; Context = $Context; Kind = $Kind }
    $g.Timer.Start()
}

function Receive-MmrGuiMessages {
    <# The lines of the background run since the last look: added to the progress; the bar shows the last state. #>
    $g = $script:Gui
    $item = $null; $progress = $null; $added = $false
    while ($g.Shared.Queue.TryDequeue([ref]$item)) {
        if ($item[0] -eq 'Progress') { $progress = $item[1]; continue }
        if ($item[0] -eq 'Step') { $progress = $null }
        Add-MmrGuiLine $item[0] $item[1] -NoScroll
        $added = $true
    }
    if ($progress) { Add-MmrGuiLine 'Progress' $progress }
    if ($added) { $g.Controls.LogScroll.ScrollToEnd() }
}

function Step-MmrGuiWork {
    <# Every 100 ms while a background run is in progress: its lines, then its end (the outcome to OnDone). #>
    $g = $script:Gui
    if (-not $g) { return }
    Receive-MmrGuiMessages
    Set-MmrGuiProgress -Tick
    $job = $g.Job
    if (-not $job -or -not $job.Handle.IsCompleted) { return }
    $g.Job = $null
    $g.Timer.Stop()
    $outcome = $null
    try {
        $output = $job.PowerShell.EndInvoke($job.Handle)
        if ($output.Count) { $outcome = $output[$output.Count - 1]; if ($null -ne $outcome) { $outcome = $outcome.psobject.BaseObject } }
        if ($outcome -isnot [hashtable]) {
            $err = @($job.PowerShell.Streams.Error) | Select-Object -First 1
            $outcome = @{ Ok = $false; Cancelled = $false; Started = $false; Error = $(if ($err) { [string]$err } else { 'The background run ended without a result.' }); Result = $null; Report = $null }
        }
    }
    catch {
        $inner = if ($_.Exception.InnerException) { $_.Exception.InnerException.Message } else { $_.Exception.Message }
        $outcome = @{ Ok = $false; Cancelled = $false; Started = $false; Error = $inner; Result = $null; Report = $null }
    }
    finally { $job.PowerShell.Dispose() }
    Receive-MmrGuiMessages
    try { & $job.OnDone $outcome $job.Context }
    catch {
        Add-MmrGuiLine 'Fail' $_.Exception.Message
        Set-MmrGuiStatus 'Failed - see the progress' 'Failed'
        if ($g.Running -and -not $g.Job) { Stop-MmrGuiRun }
    }
}

function Wait-MmrGuiWork {
    <# Lab tests and tools: waits (the window answering) until the run of the window is over, with what follows it. #>
    param([int]$TimeoutSeconds = 3600)
    $until = [datetime]::UtcNow.AddSeconds($TimeoutSeconds)
    while ($script:Gui -and ($script:Gui.Job -or $script:Gui.Running) -and [datetime]::UtcNow -lt $until) { Invoke-MmrGuiPump; Start-Sleep -Milliseconds 40 }
}

#endregion

function Invoke-MmrGuiSearch {
    $g = $script:Gui
    $c = $g.Controls
    $g.Items.Clear(); $g.Lines.Clear()
    $cfg = Get-MmrGuiSettings
    $locations = @(if ($c.LocationPrimary.IsChecked) { 'Primary' }; if ($c.LocationArchive.IsChecked) { 'Archive' })
    $formats = @(if ($c.FormatCsv.IsChecked) { 'Csv' }; if ($c.FormatHtml.IsChecked) { 'Html' })
    $layout = if ($c.LayoutPerMailbox.IsChecked) { 'PerMailbox' } elseif ($c.LayoutBoth.IsChecked) { 'Both' } else { 'Global' }
    $requestArgs = @{
        Settings = $cfg; Mailbox = @($c.Mailbox.Text); Subject = @(Split-MmrList @($c.Subject.Text)); Location = $locations; RecoverableItems = [bool]$c.RecoverableItems.IsChecked; Recipients = [bool]$c.Recipients.IsChecked
        ExcludeFolder = @(Split-MmrList @($c.ExcludeFolders.Text)); Layout = $layout; Formats = $formats
    }
    if ($c.StartDate.SelectedDate) { $requestArgs.Start = [datetime]$c.StartDate.SelectedDate }
    if ($c.EndDate.SelectedDate) { $requestArgs.End = [datetime]$c.EndDate.SelectedDate }
    $problems = [Collections.Generic.List[string]]::new()
    $request = $null
    try {
        $request = New-MmrRequest @requestArgs
        # The ArchiveGuid of a list loaded in the window.
        foreach ($e in $request.Mailboxes) { if (-not $e.ArchiveGuid -and $g.ArchiveGuids.ContainsKey($e.Address)) { $e.ArchiveGuid = $g.ArchiveGuids[$e.Address] } }
        foreach ($p in (Test-MmrRequest -Request $request).Problems) { $problems.Add($p) }
    }
    catch { $problems.Add($_.Exception.Message) }
    foreach ($p in (Test-MmrConfiguration -Configuration $cfg -ForConnection).Problems) { $problems.Add("$p (Connection, at the bottom left)") }
    if ($problems.Count) {
        foreach ($p in $problems) { Add-MmrGuiLine 'Fail' $p }
        Set-MmrGuiStatus 'Fix the values on the left' 'Failed'
        return
    }

    $secret = $null
    if ($cfg.AuthMode -eq 'ClientSecret' -and $c.Secret.SecurePassword.Length) { $secret = $c.Secret.SecurePassword.Copy() }
    Start-MmrGuiRun 'Reading...'
    $g.Result = $null; $g.LastReport = $null; $g.LastCsv = $null; $g.LastFolder = $null
    $g.MailboxRows.ReplaceAll($null); $g.PreviewRows.ReplaceAll($null); $g.Controls.PreviewInfo.Text = ''
    Write-MmrLog 'STEP' "Window search: $(@($request.Mailboxes).Count) mailbox(es), $($request.Location -join ', ')$(if ($request.RecoverableItems) { ', Recoverable Items' }), filter '$(Get-MmrMessageFilter -Start $request.Start -End $request.End -Subject $request.Subject)'"
    try {
        Start-MmrGuiWork -Kind 'Search' -Arguments @{ Settings = $cfg; Request = $request; Secret = $secret } -Context @{ Settings = $cfg; Secret = $secret } -OnDone {
            param($outcome, $context)
            $g = $script:Gui
            try {
                if ($outcome.Ok) {
                    $g.Result = $outcome.Result
                    $g.Settings = $context.Settings
                    $g.LastFolder = $outcome.Report.Directory; $g.LastReport = $outcome.Report.Html; $g.LastCsv = $outcome.Report.Csv
                    $g.Controls.Footer.Text = "Report: $($outcome.Report.Directory)"
                    Update-MmrGuiRows -Preview $outcome.Preview
                    $n = $g.Result.Counts
                    $recoverable = if ($g.Result.Request.RecoverableItems) { " $([char]0x00B7) $('{0:N0}' -f $n.RecoverableMessages) recoverable" } else { '' }
                    Set-MmrGuiStatus ('{0:N0} message(s) {1} {2:N0} primary {1} {3:N0} archive{4}' -f $n.Messages, [char]0x00B7, $n.PrimaryMessages, $n.ArchiveMessages, $recoverable) $g.Result.Status
                    Add-MmrGuiLine 'Info' 'Nothing was changed in the mailboxes. Open the report: every message, searchable, and the folders and mailboxes read.'
                }
                elseif ($outcome.Cancelled) { Add-MmrGuiLine 'Warn' 'Stopped: no report was written.'; Set-MmrGuiStatus 'Stopped' 'Warning' }
                else { Add-MmrGuiLine 'Fail' $outcome.Error; Set-MmrGuiStatus 'Failed - see the progress' 'Failed' }
            }
            finally {
                if ($context.Secret) { $context.Secret.Dispose() }
                Stop-MmrGuiRun
            }
        }
    }
    catch {
        Add-MmrGuiLine 'Fail' $_.Exception.Message
        Set-MmrGuiStatus 'Failed - see the progress' 'Failed'
        if ($secret) { $secret.Dispose() }
        Stop-MmrGuiRun
    }
}

function Show-MmrGui {
    <#
    .SYNOPSIS
        Opens the window. Default configuration: config\MailboxMessageReport.config.psd1 of the tool folder.
    #>
    [CmdletBinding()]
    param([hashtable]$Configuration)

    if (-not $Configuration) { $Configuration = Import-MmrConfiguration }
    $window = New-MmrForm -Configuration $Configuration
    # The engine of the window starts loading now, in the background: ready by the first search.
    Open-MmrGuiRunspace
    # Ctrl+C in the console would stop the command that owns the window: the window then could not run any of its
    # PowerShell handlers. Ctrl+C is ignored while the window is open.
    $previousCtrlC = $null
    try { if (-not [Console]::IsInputRedirected) { $previousCtrlC = [Console]::TreatControlCAsInput; [Console]::TreatControlCAsInput = $true } } catch { $previousCtrlC = $null }
    $previousQuiet = $script:Quiet
    try {
        # The console stays quiet: the window shows the progress (the log file still gets every line).
        $script:Quiet = $true
        [void]$window.Form.ShowDialog()
    }
    finally {
        $script:Quiet = $previousQuiet
        if ($null -ne $previousCtrlC) { try { [Console]::TreatControlCAsInput = $previousCtrlC } catch { } }
        if ($script:Gui) {
            $script:Gui.Timer.Stop()
            if ($script:Gui.Runspace) { try { $script:Gui.Runspace.Dispose() } catch { Write-MmrLog 'WARN' "Engine of the window: $($_.Exception.Message)" } }
        }
        $script:Gui = $null
    }
}
