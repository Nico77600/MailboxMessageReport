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
    Version : 2.1.1
#>

$script:Gui = $null

function Get-MmrGuiXaml {
    <# The window. Colours come from the Fluent theme resources (or the classic fallback of Set-MmrGuiTheme). #>
    @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Width="1480" Height="940" MinWidth="1080" MinHeight="680" WindowStartupLocation="CenterScreen"
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
          <RowDefinition Height="170" MinHeight="110"/>
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
            <ListView x:Name="Mailboxes" Grid.Row="1" Height="112" SelectionMode="Single" BorderThickness="0" Background="Transparent" FontSize="13"
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
        <Border Grid.Row="1" Style="{StaticResource MmrCard}" Margin="0,0,0,6" Padding="14,12,14,12">
          <Grid>
            <Grid.RowDefinitions>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="*"/>
            </Grid.RowDefinitions>
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>
              <TextBlock Text="Preview" Style="{StaticResource MmrCardTitle}"/>
              <TextBlock x:Name="PreviewInfo" Grid.Column="1" Margin="12,1,12,8" FontSize="12" VerticalAlignment="Top" TextTrimming="CharacterEllipsis" Foreground="{DynamicResource TextFillColorSecondaryBrush}"/>
              <StackPanel Grid.Column="2" Orientation="Horizontal" Margin="0,-4,0,6">
                <RadioButton x:Name="ViewFolders" GroupName="PreviewView" Content="Folders" Margin="0,0,14,0" MinWidth="0"/>
                <RadioButton x:Name="ViewList" GroupName="PreviewView" Content="List" MinWidth="0"/>
              </StackPanel>
            </Grid>
            <!-- Folders: the tree of every mailbox (primary mailbox, archive, Recoverable Items), the messages of the folder
                 selected, the message selected - like Outlook. Rows and nodes are compiled objects (PreviewRow, FolderNode). -->
            <Grid x:Name="FolderView" Grid.Row="1">
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="230" MinWidth="140"/>
                <ColumnDefinition Width="6"/>
                <ColumnDefinition Width="*" MinWidth="220"/>
                <ColumnDefinition Width="6"/>
                <ColumnDefinition Width="*" MinWidth="240"/>
              </Grid.ColumnDefinitions>
              <TreeView x:Name="FolderTree" BorderThickness="0" Background="Transparent" FontSize="13"
                        VirtualizingPanel.IsVirtualizing="True" VirtualizingPanel.VirtualizationMode="Recycling" ScrollViewer.HorizontalScrollBarVisibility="Disabled">
                <TreeView.ItemContainerStyle>
                  <Style TargetType="TreeViewItem" BasedOn="{StaticResource {x:Type TreeViewItem}}">
                    <Setter Property="IsExpanded" Value="{Binding IsExpanded, Mode=TwoWay}"/>
                    <Setter Property="IsSelected" Value="{Binding IsSelected, Mode=TwoWay}"/>
                    <Setter Property="ToolTip" Value="{Binding Tip}"/>
                    <Setter Property="Padding" Value="0,1"/>
                  </Style>
                </TreeView.ItemContainerStyle>
                <TreeView.ItemTemplate>
                  <HierarchicalDataTemplate ItemsSource="{Binding Children}">
                    <Grid>
                      <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                      </Grid.ColumnDefinitions>
                      <TextBlock Text="{Binding Glyph}" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="13" Margin="0,2,8,0" Foreground="{DynamicResource MmrBrandText}"/>
                      <TextBlock x:Name="NodeName" Grid.Column="1" Text="{Binding Name}" TextTrimming="CharacterEllipsis" Foreground="{DynamicResource TextFillColorPrimaryBrush}"/>
                      <TextBlock Grid.Column="2" Text="{Binding CountText}" Margin="8,0,0,0" FontWeight="SemiBold" Foreground="{DynamicResource MmrBrandText}"/>
                    </Grid>
                    <HierarchicalDataTemplate.Triggers>
                      <DataTrigger Binding="{Binding Dim}" Value="True">
                        <Setter TargetName="NodeName" Property="Foreground" Value="{DynamicResource TextFillColorTertiaryBrush}"/>
                      </DataTrigger>
                      <DataTrigger Binding="{Binding Status}" Value="Failed">
                        <Setter TargetName="NodeName" Property="Foreground" Value="{DynamicResource MmrCritical}"/>
                      </DataTrigger>
                    </HierarchicalDataTemplate.Triggers>
                  </HierarchicalDataTemplate>
                </TreeView.ItemTemplate>
              </TreeView>
              <GridSplitter Grid.Column="1" Width="6" HorizontalAlignment="Stretch" Background="Transparent" ResizeBehavior="PreviousAndNext"/>
              <Grid Grid.Column="2">
                <Grid.RowDefinitions>
                  <RowDefinition Height="Auto"/>
                  <RowDefinition Height="*"/>
                </Grid.RowDefinitions>
                <TextBlock x:Name="FolderTitle" FontSize="12" FontWeight="SemiBold" Margin="10,0,0,6" TextTrimming="CharacterEllipsis" Foreground="{DynamicResource TextFillColorSecondaryBrush}"/>
                <ListView x:Name="FolderMessages" Grid.Row="1" SelectionMode="Single" BorderThickness="0" Background="Transparent"
                          VirtualizingPanel.IsVirtualizing="True" VirtualizingPanel.VirtualizationMode="Recycling" ScrollViewer.HorizontalScrollBarVisibility="Disabled">
                  <ListView.ItemContainerStyle>
                    <Style TargetType="ListViewItem" BasedOn="{StaticResource {x:Type ListViewItem}}">
                      <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
                      <Setter Property="Padding" Value="10,6"/>
                    </Style>
                  </ListView.ItemContainerStyle>
                  <ListView.ItemTemplate>
                    <DataTemplate>
                      <Grid>
                        <Grid.RowDefinitions>
                          <RowDefinition Height="Auto"/>
                          <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        <Grid.ColumnDefinitions>
                          <ColumnDefinition Width="*"/>
                          <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <TextBlock Text="{Binding FromDisplay}" FontWeight="SemiBold" FontSize="13" TextTrimming="CharacterEllipsis" Foreground="{DynamicResource TextFillColorPrimaryBrush}"/>
                        <StackPanel Grid.Column="1" Orientation="Horizontal" Margin="8,0,0,0">
                          <TextBlock Text="{Binding AttachmentGlyph}" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="11" Margin="0,3,6,0" Foreground="{DynamicResource TextFillColorSecondaryBrush}"/>
                          <TextBlock Text="{Binding ReceivedShort}" FontSize="11" Margin="0,2,0,0" Foreground="{DynamicResource TextFillColorSecondaryBrush}"/>
                        </StackPanel>
                        <TextBlock Grid.Row="1" Grid.ColumnSpan="2" Text="{Binding Subject}" FontSize="12" TextTrimming="CharacterEllipsis" Foreground="{DynamicResource MmrBrandText}"/>
                      </Grid>
                    </DataTemplate>
                  </ListView.ItemTemplate>
                </ListView>
              </Grid>
              <GridSplitter Grid.Column="3" Width="6" HorizontalAlignment="Stretch" Background="Transparent" ResizeBehavior="PreviousAndNext"/>
              <Border x:Name="ReadingPane" Grid.Column="4" BorderThickness="1,0,0,0" BorderBrush="{DynamicResource DividerStrokeColorDefaultBrush}" Padding="14,0,0,0">
                <Grid>
                  <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                  </Grid.RowDefinitions>
                  <StackPanel x:Name="ReadingHeader">
                    <TextBlock x:Name="ReadSubject" FontSize="16" FontWeight="SemiBold" TextWrapping="Wrap" Foreground="{DynamicResource TextFillColorPrimaryBrush}"/>
                    <TextBlock x:Name="ReadFrom" FontSize="13" FontWeight="SemiBold" Margin="0,8,0,0" TextWrapping="Wrap" Foreground="{DynamicResource TextFillColorPrimaryBrush}"/>
                    <TextBlock x:Name="ReadDetails" FontSize="12" Margin="0,4,0,0" TextWrapping="Wrap" Foreground="{DynamicResource TextFillColorSecondaryBrush}"/>
                    <Border Height="1" Margin="0,10,0,8" Background="{DynamicResource DividerStrokeColorDefaultBrush}"/>
                  </StackPanel>
                  <TextBox x:Name="ReadBody" Grid.Row="1" IsReadOnly="True" TextWrapping="Wrap" BorderThickness="0" Background="Transparent" Padding="0" FontSize="13"
                           VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled" Foreground="{DynamicResource TextFillColorPrimaryBrush}"/>
                  <TextBlock x:Name="ReadEmpty" Grid.RowSpan="2" Margin="0,40,0,0" TextWrapping="Wrap" HorizontalAlignment="Center" TextAlignment="Center" FontSize="13" Foreground="{DynamicResource TextFillColorTertiaryBrush}"
                             Text="Select a message to read it."/>
                </Grid>
              </Border>
            </Grid>
            <!-- List: every message of the preview, flat and sortable. -->
            <ListView x:Name="Preview" Grid.Row="1" SelectionMode="Single" BorderThickness="0" Background="Transparent" FontSize="13" Visibility="Collapsed"
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
                       Text="Type the mailboxes, the period or the subjects, then Read the messages. The window shows the first messages of each folder; the report holds them all."/>
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
        'Mailboxes', 'Preview', 'PreviewInfo', 'PreviewEmpty', 'ViewFolders', 'ViewList', 'FolderView', 'FolderTree', 'FolderTitle', 'FolderMessages', 'ReadingPane', 'ReadingHeader', 'ReadSubject', 'ReadFrom', 'ReadDetails', 'ReadBody', 'ReadEmpty', 'ProgressBar', 'ProgressText', 'ProgressInfo', 'LogScroll', 'Log', 'Actions', 'Search', 'Stop', 'Footer',
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
    $treeRows = [MailboxMessageReportNative.BulkCollection]::new()
    $folderRows = [MailboxMessageReportNative.BulkCollection]::new()
    $items = [Collections.ObjectModel.ObservableCollection[object]]::new()
    $controls.Mailboxes.ItemsSource = $mailboxRows
    $controls.Preview.ItemsSource = $previewRows
    $controls.FolderTree.ItemsSource = $treeRows
    $controls.FolderMessages.ItemsSource = $folderRows
    $controls.ViewFolders.IsChecked = $true
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
        ArchiveGuids = @{}; MailboxRows = $mailboxRows; PreviewRows = $previewRows; TreeRows = $treeRows; FolderRows = $folderRows; Items = $items; Lines = [Collections.Generic.List[string]]::new()
        # The reading pane: the message shown, the one waiting for its content (read a moment after it is selected).
        Reading = $null; BodyWanted = $null; BodyTimer = $null; CanReadBody = $false; MailboxKeys = @{}
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
    $controls.ViewFolders.Add_Checked({ Update-MmrGuiView })
    $controls.ViewList.Add_Checked({ Update-MmrGuiView })
    $controls.FolderTree.Add_SelectedItemChanged({ Show-MmrGuiFolder })
    $controls.FolderMessages.Add_SelectionChanged({ Show-MmrGuiMessage $script:Gui.Controls.FolderMessages.SelectedItem })
    # The content of a message is read a moment after it is selected: moving through the list reads only where it stops.
    $bodyTimer = [Windows.Threading.DispatcherTimer]::new([Windows.Threading.DispatcherPriority]::Background)
    $bodyTimer.Interval = [TimeSpan]::FromMilliseconds(300)
    $bodyTimer.Add_Tick({ $script:Gui.BodyTimer.Stop(); Read-MmrGuiBody })
    $script:Gui.BodyTimer = $bodyTimer
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
    else { "One file for every mailbox. The HTML report holds every message (up to $('{0:N0}' -f [int]$g.Configuration.HtmlMaxMessages)), with filters; the CSV file as well." }
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
    <# The mailboxes, the folder tree and the preview from the result of the run, replaced in one go. #>
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
    $messages = [MailboxMessageReportNative.PreviewRow]::Build($Preview)
    $g.PreviewRows.ReplaceAll($messages)
    # The folder view: every mailbox read, its primary mailbox and archive, their folders; the messages in their folder.
    $g.MailboxKeys = @{}
    foreach ($m in @(if ($g.Result) { $g.Result.Mailboxes })) { $g.MailboxKeys["$($m.Address)|Primary"] = $m.Key; if ($m.ArchiveKey) { $g.MailboxKeys["$($m.Address)|Archive"] = $m.ArchiveKey } }
    # Assigned directly: an 'if' would unroll a list of one mailbox into its node.
    $roots = $null
    if ($g.Result) { $roots = [MailboxMessageReportNative.FolderNode]::Build($g.Result.Mailboxes, $g.Result.Folders, $messages) }
    $g.TreeRows.ReplaceAll($roots)
    $g.FolderRows.ReplaceAll($null)
    Show-MmrGuiMessage $null
    $many = @($rows).Count -gt 1
    foreach ($column in $g.Controls.Preview.View.Columns) { if ($column.Header -eq 'Mailbox') { $column.Width = if ($many) { 180 } else { 0 } } }
    Update-MmrGuiColumns
    $total = if ($g.Result) { [long]$g.Result.Counts.Messages } else { 0 }
    $shown = $g.PreviewRows.Count
    $perFolder = [int]$g.Configuration.PreviewPerFolder
    $g.Controls.PreviewInfo.Text = if (-not $g.Result) { '' }
    elseif ($total -gt $shown) { "$('{0:N0}' -f $shown) of $('{0:N0}' -f $total) messages$(if ($perFolder -gt 0) { ", the first $perFolder of each folder" }) $($script:Dot) every message is in the report" }
    else { "$('{0:N0}' -f $total) message(s)" }
    if ($g.Result) { $n = $g.Result.Counts; $g.Controls.Counts.Text = '{0:N0} read {1} {2:N0} with an archive {1} {3:N0} not read' -f ($n.MailboxesRead + $n.MailboxesPartial), $script:Dot, $n.WithArchive, $n.MailboxesNotRead }
    else { $g.Controls.Counts.Text = '' }
    # The first folder with messages, selected (its parents open).
    if ($roots) {
        $first = Select-MmrGuiFirstFolder $roots
        # Selected in the tree once it is drawn; its messages shown now.
        if ($first) { $first.IsSelected = $true; Show-MmrGuiFolder $first }
    }
    Update-MmrGuiView
}

function Select-MmrGuiFirstFolder {
    <# The first folder of the tree with messages in the preview, its parents opened. #>
    param($Nodes)
    foreach ($node in @($Nodes)) {
        if ($node.Kind -eq 'Folder' -and $node.Messages.Count) { return $node }
        $found = Select-MmrGuiFirstFolder $node.Children
        if ($found) { $node.IsExpanded = $true; return $found }
    }
    return $null
}

function Update-MmrGuiView {
    <# Folders (tree, messages, reading pane) or List (every message of the preview), and the empty text. #>
    $g = $script:Gui
    if (-not $g) { return }
    $c = $g.Controls
    $any = $g.PreviewRows.Count -gt 0 -or $g.TreeRows.Count -gt 0
    $folders = [bool]$c.ViewFolders.IsChecked
    $c.FolderView.Visibility = if ($any -and $folders) { 'Visible' } else { 'Collapsed' }
    $c.Preview.Visibility = if ($any -and -not $folders) { 'Visible' } else { 'Collapsed' }
    $c.PreviewEmpty.Visibility = if ($any) { 'Collapsed' } else { 'Visible' }
    if (-not $any -and $g.Result) { $c.PreviewEmpty.Text = 'No message: widen the period, check the subjects, or tick the archive.' }
    Update-MmrGuiColumns
}

function Show-MmrGuiFolder {
    <# The messages of the folder selected in the tree (the first ones of the folder: the preview), newest first. #>
    param($Node)
    $g = $script:Gui
    if (-not $g) { return }
    $node = if ($Node) { $Node } else { $g.Controls.FolderTree.SelectedItem }
    if (-not $node -or $node.Kind -ne 'Folder') {
        $g.FolderRows.ReplaceAll($null)
        $g.Controls.FolderTitle.Text = if ($node) { "$($node.Name) $($script:Dot) $('{0:N0}' -f $node.Found) message(s) found $($script:Dot) select a folder" } else { '' }
        Show-MmrGuiMessage $null
        return
    }
    $g.FolderRows.ReplaceAll($node.Messages)
    $shown = $node.Messages.Count
    $g.Controls.FolderTitle.Text = if ($node.Found -gt $shown) { "$($node.Name) $($script:Dot) the first $shown of $('{0:N0}' -f $node.Found) messages" }
    elseif ($node.Found) { "$($node.Name) $($script:Dot) $('{0:N0}' -f $node.Found) message(s)" }
    elseif ($node.Status -eq 'Failed') { "$($node.Name) $($script:Dot) not read" }
    else { "$($node.Name) $($script:Dot) no message$(if ($node.Items) { " found ($('{0:N0}' -f $node.Items) item(s) in the folder)" })" }
    if ($shown) { $g.Controls.FolderMessages.SelectedIndex = 0 } else { Show-MmrGuiMessage $null }
}

function Show-MmrGuiMessage {
    <# The reading pane: the message selected - who, to whom, when, where - and its content, read on demand. #>
    param($Row)
    $g = $script:Gui
    if (-not $g) { return }
    $c = $g.Controls
    $g.Reading = $Row
    if (-not $Row) {
        $c.ReadingHeader.Visibility = 'Collapsed'; $c.ReadBody.Visibility = 'Collapsed'; $c.ReadEmpty.Visibility = 'Visible'
        $c.ReadEmpty.Text = if ($g.FolderRows.Count) { 'Select a message to read it.' } else { 'Select a folder, then a message.' }
        return
    }
    $dot = $script:Dot
    $c.ReadingHeader.Visibility = 'Visible'; $c.ReadBody.Visibility = 'Visible'; $c.ReadEmpty.Visibility = 'Collapsed'
    $c.ReadSubject.Text = $Row.Subject
    $from = if ($Row.From -and $Row.FromDisplay -ne $Row.From) { "$($Row.FromDisplay) <$($Row.From)>" } else { $Row.FromDisplay }
    if ($Row.Sender -and $Row.From -and $Row.Sender -ne $Row.From) { $from += "  (sent by $($Row.Sender))" }
    $c.ReadFrom.Text = $from
    $lines = [Collections.Generic.List[string]]::new()
    $lines.Add("To: $(if ($Row.To) { $Row.To } else { '-' })")
    if ($Row.Cc) { $lines.Add("Cc: $($Row.Cc)") }
    if ($Row.Bcc) { $lines.Add("Bcc: $($Row.Bcc)") }
    $lines.Add("Received $($Row.Received)$(if ($Row.Sent) { " $dot sent $($Row.Sent)" })")
    $lines.Add("$($Row.LocationText) $dot $($Row.FolderPath) $dot $($Row.Mailbox)")
    $extra = @(if ($Row.HasAttachments) { 'Attachments' }; if ($Row.Importance -and $Row.Importance -ne 'Normal') { "Importance $($Row.Importance)" }; if ($Row.Type -and $Row.Type -ne 'Message') { $Row.Type })
    if ($extra.Count) { $lines.Add($extra -join " $dot ") }
    if ($Row.InternetMessageId) { $lines.Add("Message ID: $($Row.InternetMessageId)") }
    $c.ReadDetails.Text = $lines -join [Environment]::NewLine
    $c.ReadBody.ScrollToHome()
    if ($Row.BodyLoaded) { $c.ReadBody.Text = $Row.Body; return }
    if (-not $g.Configuration.ReadBody) { $c.ReadBody.Text = 'The window does not read the content of the messages (Window.ReadBody).'; return }
    if (-not $g.CanReadBody) { $c.ReadBody.Text = 'The content needs the application permission Mail.Read: this application reads the subject, the sender, the recipients and the dates only (Mail.ReadBasic.All).'; return }
    $c.ReadBody.Text = 'Reading the content...'
    $g.BodyWanted = $Row
    $g.BodyTimer.Stop(); $g.BodyTimer.Start()
}

function Read-MmrGuiBody {
    <# The content of the message waiting in the reading pane, read in the background runspace (its Graph connection). #>
    $g = $script:Gui
    if (-not $g -or -not $g.BodyWanted) { return }
    # A run, or another content being read: tried again in a moment.
    if ($g.Running -or $g.Job) { $g.BodyTimer.Start(); return }
    $row = $g.BodyWanted
    $g.BodyWanted = $null
    $key = $g.MailboxKeys["$($row.Mailbox)|$($row.Location -replace ' \(RI\)$', '')"]
    if (-not $key -or -not $row.ItemId) { $row.Body = 'The content cannot be read: mailbox or message not known.'; $row.BodyLoaded = $true; if ($g.Reading -eq $row) { $g.Controls.ReadBody.Text = $row.Body }; return }
    $url = "$(Get-MmrUserPath $key)/messages/$([Uri]::EscapeDataString($row.ItemId))?`$select=body"
    try {
        Start-MmrGuiWork -Kind 'Body' -Arguments @{ Url = $url; Mailbox = $row.Mailbox; Where = "$($row.LocationText) $($row.FolderPath), message ID $($row.InternetMessageId)" } -Context @{ Row = $row } -OnDone {
            param($outcome, $context)
            $g = $script:Gui
            $r = $context.Row
            $r.Body = if ($outcome.Ok) { $outcome.Body } else { "The content could not be read: $($outcome.Error)" }
            $r.BodyLoaded = [bool]$outcome.Ok
            if ($g.Reading -eq $r) { $g.Controls.ReadBody.Text = $r.Body }
        }
    }
    catch { $row.Body = "The content could not be read: $($_.Exception.Message)"; if ($g.Reading -eq $row) { $g.Controls.ReadBody.Text = $row.Body } }
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
        The work of the window, in its background runspace (or inline). Search = connection, mailboxes, folders,
        messages, report. Body = the content of one message, as text, with the connection of the last search (each read
        written to the log). Never throws: returns @{ Ok; Cancelled; Error; Result; Report; Preview; Body }.
    #>
    param([Parameter(Mandatory = $true)][string]$Kind, [hashtable]$Arguments = @{}, [Parameter(Mandatory = $true)][hashtable]$Shared)
    $script:Ui = $Shared
    $script:Quiet = $true
    if ($Shared.Log -and -not [object]::ReferenceEquals($script:LogWriter, $Shared.Log)) { $script:LogWriter = $Shared.Log }
    $a = $Arguments
    $out = @{ Ok = $false; Cancelled = $false; Error = ''; Result = $null; Report = $null; Preview = $null; Body = '' }
    $partsPath = $null
    try {
        if ($Kind -eq 'Body') {
            if (-not $script:Graph) { throw 'Not connected: read the messages again.' }
            $r = Invoke-MmrGraph -Path $a.Url -Headers @{ Prefer = 'outlook.body-content-type="text"' }
            if ($r.Status -ne 200) { throw "$($r.Status) $($r.ErrorCode): $($r.ErrorMessage)" }
            Write-MmrLog 'INFO' "Window: content of a message read - $($a.Mailbox), $($a.Where)"
            $text = [string](Get-MmrProperty (Get-MmrProperty $r.Body 'body') 'content')
            # Very long contents are cut: the reading pane is a preview.
            $out.Body = if ($text.Length -gt 200000) { $text.Substring(0, 200000) + "`n`n[...] (cut at 200,000 characters)" } elseif ($text.Trim()) { $text.Trim() } else { '(no content)' }
            $out.Ok = $true
            return $out
        }
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
        $report = Export-MmrReport -Result $out.Result -Directory $runPath -Prefix $s.ReportPrefix -Formats $a.Request.Formats -Layout $a.Request.Layout -Delimiter $s.CsvDelimiter -HtmlMaxMessages $s.HtmlMaxMessages -PreviewMessages $s.PreviewMessages -PreviewPerFolder $s.PreviewPerFolder -PartsPath $partsPath
        Write-MmrItem Ok "Report: $($report.Directory)" -Icon File
        $csv = if ($report.Files.Contains('Messages')) { $report.Files.Messages } elseif ($report.Files.Contains('Mailboxes')) { $report.Files.Mailboxes } else { $null }
        $out.Report = @{ Directory = $report.Directory; Html = Get-MmrProperty $report.Files 'Html'; Csv = $csv }
        $out.Preview = $report.Preview
        $out.CanReadBody = @($connection.Roles | Where-Object { $_ -in 'Mail.Read', 'Mail.ReadWrite' }).Count -gt 0
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
    # The content of a message still being read: finished first (a second or so).
    $until = [datetime]::UtcNow.AddSeconds(30)
    while ($g.Job -and [datetime]::UtcNow -lt $until) { Invoke-MmrGuiPump; Start-Sleep -Milliseconds 30 }
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
    $g.MailboxRows.ReplaceAll($null); $g.PreviewRows.ReplaceAll($null); $g.TreeRows.ReplaceAll($null); $g.FolderRows.ReplaceAll($null); $g.Controls.PreviewInfo.Text = ''
    $g.BodyWanted = $null; $g.BodyTimer.Stop(); Show-MmrGuiMessage $null; Update-MmrGuiView
    Write-MmrLog 'STEP' "Window search: $(@($request.Mailboxes).Count) mailbox(es), $($request.Location -join ', ')$(if ($request.RecoverableItems) { ', Recoverable Items' }), filter '$(Get-MmrMessageFilter -Start $request.Start -End $request.End -Subject $request.Subject)', recipients $(if (Get-MmrProperty $request 'Recipients') { 'yes' } else { 'no' }), report $(@($request.Formats) -join '+') $($request.Layout)"
    try {
        Start-MmrGuiWork -Kind 'Search' -Arguments @{ Settings = $cfg; Request = $request; Secret = $secret } -Context @{ Settings = $cfg; Secret = $secret } -OnDone {
            param($outcome, $context)
            $g = $script:Gui
            try {
                if ($outcome.Ok) {
                    $g.Result = $outcome.Result
                    $g.Settings = $context.Settings
                    $g.LastFolder = $outcome.Report.Directory; $g.LastReport = $outcome.Report.Html; $g.LastCsv = $outcome.Report.Csv
                    $g.CanReadBody = [bool]$outcome['CanReadBody']
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

# SIG # Begin signature block
# MIImIQYJKoZIhvcNAQcCoIImEjCCJg4CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCziWtlOSrIdQWz
# t1KlBmOLPu1s13t9RRAF6upRWACQnqCCDW8wggZyMIIEWqADAgECAghkM1HTxzif
# CDANBgkqhkiG9w0BAQsFADB8MQswCQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMx
# EDAOBgNVBAcMB0hvdXN0b24xGDAWBgNVBAoMD1NTTCBDb3Jwb3JhdGlvbjExMC8G
# A1UEAwwoU1NMLmNvbSBSb290IENlcnRpZmljYXRpb24gQXV0aG9yaXR5IFJTQTAe
# Fw0xNjA2MjQyMDQ0MzBaFw0zMTA2MjQyMDQ0MzBaMHgxCzAJBgNVBAYTAlVTMQ4w
# DAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91c3RvbjERMA8GA1UECgwIU1NMIENv
# cnAxNDAyBgNVBAMMK1NTTC5jb20gQ29kZSBTaWduaW5nIEludGVybWVkaWF0ZSBD
# QSBSU0EgUjEwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQCfgxNzqrDG
# bSHL24t6h3TQcdyOl3Ka5LuINLTdgAPGL0WkdJq/Hg9Q6p5tePOf+lEmqT2d0bKU
# Vz77OYkbkStW72fL5gvjDjmMxjX0jD3dJekBrBdCfVgWQNz51ShEHZVkMGE6ZPKX
# 13NMfXsjAm3zdetVPW+qLcSvvnSsXf5qtvzqXHnpD0OctVIFD+8+sbGP0EmtpuNC
# GVQ/8y8Ooct8/hP5IznaJRy4PgBKOm8yMDdkHseudQfYVdIYyQ6KvKNc8HwKp4WB
# wg6vj5lc02AlvINaaRwlE81y9eucgJvcLGfE3ckJmNVz68Qho+Uyjj4vUpjGYDdk
# jLJvSlRyGMwnh/rNdaJjIUy1PWT9K6abVa8mTGC0uVz+q0O9rdATZlAfC9KJpv/X
# gAbxwxECMzNhF/dWH44vO2jnFfF3VkopngPawismYTJboFblSSmNNqf1x1KiVgMg
# Lzh4gL32Bq5BNMuURb2bx4kYHwu6/6muakCZE93vUN8BuvIE1tAx3zQ4XldbyDge
# VtSsSKbt//m4wTvtwiS+RGCnd83VPZhZtEPqqmB9zcLlL/Hr9dQg1Zc0bl0EawUR
# 0tOSjAknRO1PNTFGfnQZBWLsiePqI3CY5NEv1IoTGEaTZeVYc9NMPSd6Ij/D+KNV
# t/nmh4LsRR7Fbjp8sU65q2j3m2PVkUG8qQIDAQABo4H7MIH4MA8GA1UdEwEB/wQF
# MAMBAf8wHwYDVR0jBBgwFoAU3QQJB6L1en1SUxKSle44gCUNplkwMAYIKwYBBQUH
# AQEEJDAiMCAGCCsGAQUFBzABhhRodHRwOi8vb2NzcHMuc3NsLmNvbTARBgNVHSAE
# CjAIMAYGBFUdIAAwEwYDVR0lBAwwCgYIKwYBBQUHAwMwOwYDVR0fBDQwMjAwoC6g
# LIYqaHR0cDovL2NybHMuc3NsLmNvbS9zc2wuY29tLXJzYS1Sb290Q0EuY3JsMB0G
# A1UdDgQWBBRUwv4QlQCTzWr158DX2bJLuI8M4zAOBgNVHQ8BAf8EBAMCAYYwDQYJ
# KoZIhvcNAQELBQADggIBAPUPJodwr5miyvXWyfCNZj05gtOII9iCv49UhCe204MH
# 154niU2EjlTRIO5gQ9tXQjzHsJX2vszqoz2OTwbGK1mGf+tzG8rlQCbgPW/M9r1x
# xs19DiBAOdYF0q+UCL9/wlG3K7V7gyHwY9rlnOFpLnUdTsthHvWlM98CnRXZ7WmT
# V7pGRS6AvGW+5xI+3kf/kJwQrfZWsqTU+tb8LryXIbN2g9KR+gZQ0bGAKID+260P
# Z+34fdzZcFt6umi1s0pmF4/n8OdX3Wn+vF7h1YyfE7uVmhX7eSuF1W0+Z0duGwdc
# +1RFDxYRLhHDsLy1bhwzV5Qe/kI0Ro4xUE7bM1eV+jjk5hLbq1guRbfZIsr0WkdJ
# LCjoT4xCPGRo6eZDrBmRqccTgl/8cQo3t51Qezxd96JSgjXktefTCm9r/o35pNfV
# HUvnfWII+NnXrJlJ27WEQRQu9i5gl1NLmv7xiHp0up516eDap8nMLDt7TAp4z5T3
# NmC2gzyKVMtODWgqlBF1JhTqIDfM63kXdlV4cW3iSTgzN9vkbFnHI2LmvM4uVEv9
# XgMqyN0eS3FE0HU+MWJliymm7STheh2ENH+kF3y0rH0/NVjLw78a3Z9UVm1F5VPz
# iIorMaPKPlDRADTsJwjDZ8Zc6Gi/zy4WZbg8Zv87spWrmo2dzJTw7XhQf+xkR6Od
# MIIG9TCCBN2gAwIBAgIQBof4ljXihJpmCqqt2mhAWzANBgkqhkiG9w0BAQsFADB4
# MQswCQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNVBAcMB0hvdXN0b24x
# ETAPBgNVBAoMCFNTTCBDb3JwMTQwMgYDVQQDDCtTU0wuY29tIENvZGUgU2lnbmlu
# ZyBJbnRlcm1lZGlhdGUgQ0EgUlNBIFIxMB4XDTI2MTAxMzAyNDgwMloXDTI3MTAx
# MzAyNDgwMlowXDELMAkGA1UEBhMCRlIxGzAZBgNVBAcMEkNoYW50ZWxvdXAtRW4t
# QnJpZTEXMBUGA1UECgwOTmljb2xhcyBGYWJlcnQxFzAVBgNVBAMMDk5pY29sYXMg
# RmFiZXJ0MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAsg6CT0mFQ4jB
# V2plIS4uWFSoVqRHd82Y7XYZ4CUTe/7nPBRUlXXdKTZMMHzGgiG8dEXNE8K1Xe90
# Dgl/FnorzAFlrNWSwf6t27UAUw9xkqtRDVVpVgpixpWkAwjcQeLG+YMqBkGbFi35
# TGvS23lPFWvLp4DxARrv15MtZ/NjNTCAAGUlMP2hFVBknEmrnYjEGyeJ93sl20RS
# WXQQLNqlenBdkGCrQW4XjHtRjJAEqAC68hN6nWevopSQ8CLUV7A1RFfHFarDrzrL
# zr0/NjIdSyIfM9hShln31NDrktB6Ups2XIyrtPDNB7fjkW4jpLlOqQGfYKsU9ErE
# 4SnFsh0yOOjitegyh1RvZ9wH++3DF4CmJfTli5l/gEIm8xCCryxI+K5EogW3wu1y
# u8NT2hcvHjdh7OpRyPyFAuLZ8MfX7kPAJpnuCeOuktzY/20eVD8bDglj+4zgqzsn
# cR/t6T/7w1/HxpRyrVoKvEyRx+mphYh5jMdEa0BjxOL0vAJpdGZSfXGmPm/yPqRn
# 5STTkMtzkVQXPbcvgKM0luzFcQKPHk3qdHLEXa8gf93ktYDbLE5Wd2UHrXz3SZxN
# URHPdIXCr1qgrqQve7041kMpzrkDI5DnG2lvRG4108gp02bHG0muKyZlDdCOra7h
# Zonu2nO79sFallruzzNi0o994PoWksUCAwEAAaOCAZUwggGRMAwGA1UdEwEB/wQC
# MAAwHwYDVR0jBBgwFoAUVML+EJUAk81q9efA19myS7iPDOMwegYIKwYBBQUHAQEE
# bjBsMEgGCCsGAQUFBzAChjxodHRwOi8vY2VydC5zc2wuY29tL1NTTGNvbS1TdWJD
# QS1Db2RlU2lnbmluZy1SU0EtNDA5Ni1SMS5jZXIwIAYIKwYBBQUHMAGGFGh0dHA6
# Ly9vY3Nwcy5zc2wuY29tMFEGA1UdIARKMEgwCAYGZ4EMAQQBMDwGDCsGAQQBgqkw
# AQMDATAsMCoGCCsGAQUFBwIBFh5odHRwczovL3d3dy5zc2wuY29tL3JlcG9zaXRv
# cnkwEwYDVR0lBAwwCgYIKwYBBQUHAwMwTQYDVR0fBEYwRDBCoECgPoY8aHR0cDov
# L2NybHMuc3NsLmNvbS9TU0xjb20tU3ViQ0EtQ29kZVNpZ25pbmctUlNBLTQwOTYt
# UjEuY3JsMB0GA1UdDgQWBBRoUx3630l9wA1fUpkKkw2GUIxl/jAOBgNVHQ8BAf8E
# BAMCB4AwDQYJKoZIhvcNAQELBQADggIBAE3oj6cNuwlDToi1BwjgsA5djfRN4DIi
# BYDFZIRYaIRe/rEdnItz3aqQQdFuwXnJA/MQEmi2WAGek8GeZF+R41s8FbcDxBUE
# BoSt98TuBE+0xgxQWDiPhhXaX1Aw9V0GF4ZnlfCSw1eu2KiASaKn7Bo0PLbEAXff
# vftRXof/kyqtw8iID6I4Rhrf5DI3WVZQodakUUmWnbnziWvEZOWbNOfzwP+7BwEt
# 7jQANfbmEfb4p75rfy5dQVwniTDSvdo6xkrbe1Cr93GS7cJta18IG5CoIVUmdK+s
# W38B1sLuy8Mi0MstMXV3004yaRuzLdNaAM0f9WAFerxemiQC0kdmA44/lWV3lVkM
# WJSwkHBkkSg8Eh0hxcGzTH2xmtDhLJHNNLRwkI9u2Ur+HbLCt50oBgSV8BcCs/1i
# qi884iC/yuX+HIv9YgTUNZ8c4lLaq4VuliJpXLu2buHIsTkO0v0+5oCM91zjTVZm
# ULeH38N0Vr83psiJVuBSxQyt5nRVrpYz7FACq5gtXELkyViq/rGOlJKaRrgQ6WWq
# eoIXqKhVFAdP/jAEm2M8ezoVYsUgPyKY2OcuecV9coulAcxUIUydx6qYiOfsGj5u
# Dy7l4aYjvB771wkUMzUQwxzcDYXJtMMTaWRe6n9EemR67jc73u5CJyFg3PfA5WR1
# wZislcdCZuS6MYIYCDCCGAQCAQEwgYwweDELMAkGA1UEBhMCVVMxDjAMBgNVBAgM
# BVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMREwDwYDVQQKDAhTU0wgQ29ycDE0MDIG
# A1UEAwwrU1NMLmNvbSBDb2RlIFNpZ25pbmcgSW50ZXJtZWRpYXRlIENBIFJTQSBS
# MQIQBof4ljXihJpmCqqt2mhAWzANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3
# AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisG
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDx3rrqno88
# l3GWugSlGY7DoHIoqIC6gE23QpkUvvye9TANBgkqhkiG9w0BAQEFAASCAgAEFcGq
# qPgBMIDcMsgZt/jrojXJmrTn+lv45Wc2Xv+AUgQSS6pkGiey3ndZdJil3KP11yQH
# gNw8HhmvPQ7jgnYjwWOBlscaCI8JtMpOnXFRPZBZXV7LYEwBboqWCoIEvFlGkD9Z
# 8lRLWIVBQjeNf+YBzUg5j33UkaNpfosE7Q/QwIdKrAVG+XtQOuZ2JMWmLBux5HPN
# TNS/UAEXzrinMYEj8SB49ZfwwBZvPVEQslda/Cwt+UAfUgGoAng5Pgqso1yMjwtK
# fi8nJ80btJO2SIjp08ziaG3Mjw2vibMvSC8TVLNVE06ofrahiPyBqO6dkgXoQTun
# 9me2CkTlKWTrmwZxbKpmsOKeASn+GTd67PY2Mr5l5/8YUXgqkYJoLZl0Jk0Q1OAY
# dOo9ZWZ/akcm3fulXJ3/SRoV+5noOTCrYv2nkEasqWQD0d2fy+rzGr95YoUDKy93
# c+C8ylWq6WWfyCByxxhXwKOS2ZyP/Fgqwx4hsa2uJrwHk8gGpzDKD8rpxvlGi2bY
# ZYYozfj7TA4Cw24FdpMRZ2POpPNxIQoJ+Bo2l7IVkpPRpP3/kguhT86lFra+Gsjv
# YYB2OVFs0YXC1exa8kamYJVe6m1sO2IxngToadH3pirn9LUhBDyeR/93h14awvT8
# X2HXg4tkibCIB6JYN/4UXGvkkuZOAxF9aYmBgaGCFMUwghTBBgorBgEEAYI3AwMB
# MYIUsTCCFK0GCSqGSIb3DQEHAqCCFJ4wghSaAgEDMQ0wCwYJYIZIAWUDBAIBMIGB
# BgsqhkiG9w0BCRABBKByBHAwbgIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEIAJFy3vogPG9OGPNfkirTdTc6lkFvlfiIW8YqekCVl5BAggU8AvckmeA
# /BgPMjAyNjExMDUxNzUxMjJaMAMCAQECCCTC8vcAO072oIIRpzCCBPMwggLboAMC
# AQICEFRD+De9Wu4YAYRyuFzajuMwDQYJKoZIhvcNAQELBQAwcjELMAkGA1UEBhMC
# VVMxDjAMBgNVBAgMBVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMREwDwYDVQQKDAhT
# U0wgQ29ycDEuMCwGA1UEAwwlU1NMLmNvbSBUaW1lc3RhbXBpbmcgSXNzdWluZyBS
# U0EgQ0EgMjAeFw0yNjA2MTIxNzIyNDZaFw0zNjA2MDkxNzIyNDVaMG4xCzAJBgNV
# BAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91c3RvbjERMA8GA1UE
# CgwIU1NMIENvcnAxKjAoBgNVBAMMIVNTTC5jb20gVGltZXN0YW1waW5nIFVuaXQg
# MjAyNiBFMTBZMBMGByqGSM49AgEGCCqGSM49AwEHA0IABOgli7hmEy5IkjjzCzbJ
# DgbYV7CIUUgOOn3jkML0i3A2jYer8VMwShMdSnRQURbOzh65x/ytFQOjGw0CBaVI
# ND2jggFSMIIBTjAfBgNVHSMEGDAWgBRwmmXBQ4Cj/OmxZIlu8TCcCSHJfDBNBggr
# BgEFBQcBAQRBMD8wPQYIKwYBBQUHMAKGMWh0dHA6Ly9jZXJ0LnNzbC5jb20vU1NM
# LmNvbS10aW1lU3RhbXBpbmctSS1SMi5jZXIwUQYDVR0gBEowSDA8BgwrBgEEAYKp
# MAEDBgEwLDAqBggrBgEFBQcCARYeaHR0cHM6Ly93d3cuc3NsLmNvbS9yZXBvc2l0
# b3J5MAgGBmeBDAEEAjAWBgNVHSUBAf8EDDAKBggrBgEFBQcDCDBCBgNVHR8EOzA5
# MDegNaAzhjFodHRwOi8vY3Jscy5zc2wuY29tL1NTTC5jb20tdGltZVN0YW1waW5n
# LUktUjIuY3JsMB0GA1UdDgQWBBQnrANWuihquhc7gz1Af3moJOzDPDAOBgNVHQ8B
# Af8EBAMCB4AwDQYJKoZIhvcNAQELBQADggIBAE92bs78A57CGU1wZ/oY5rgV4Mwd
# dn22MU/bmHumb4b6W2VinMg0IGnBE5D8PJOsnqp2AmMhWdOt5TDxKgIQA9pzBute
# GYez7R1vQ1qn7IIr8Wz2AXjfre1MYJGE93saW1jHTxKgbVlchzry3CrzMYCVnyYw
# DcFZMFQ35Mv4+3H9vPoopjfDynO0v7vPy3p0UZ1dltmH53EZ5M/Qn7OZ0Q82mYLo
# h666dhscIJzUEh+BgYxzpdz8y7J1WxqGSouGRI4F7qbOJhhjiu/hgdhwJUJ/jkve
# OHB6pZB9IBJbdlns6AjUaRC6YfamsbgZ40hKl1ln4BOshgcDtLsSjaqVlUx5gMK+
# krwT6H8bm1yLhJ/Yr0eAizrgMVEqyeNWQYP6qhtKCYxehpESm8au3MlkpBGWCYZi
# goNh4OX1x2oDsYsG5p+ST2db4Vgq25N7/3U8doL0yufnqwuR/f603JY1CO++sndq
# 34Y35Kb8gLG7V7OA+5658PIvcKK+fn3bUUiQVcasy5spBX7aC9FKdkPvrciglHa4
# m3XJ6weBykt7PrdwI3XGwNOgwi/Bon0IvXy9MkQUTZQOKOgT/kq++eTGgjY1EWGQ
# 7w8sSRwQU+3Ko2SMLIkfLPuvOtJ6BWwQbXYHtiMOwjwv9UN9SNu4VaJUuwdMmqN7
# /F9jLVst5rGZV03BMIIF3TCCA8WgAwIBAgIIeyyb0xaAMpkwDQYJKoZIhvcNAQEL
# BQAwfDELMAkGA1UEBhMCVVMxDjAMBgNVBAgMBVRleGFzMRAwDgYDVQQHDAdIb3Vz
# dG9uMRgwFgYDVQQKDA9TU0wgQ29ycG9yYXRpb24xMTAvBgNVBAMMKFNTTC5jb20g
# Um9vdCBDZXJ0aWZpY2F0aW9uIEF1dGhvcml0eSBSU0EwHhcNMTYwMjEyMTczOTM5
# WhcNNDEwMjEyMTczOTM5WjB8MQswCQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMx
# EDAOBgNVBAcMB0hvdXN0b24xGDAWBgNVBAoMD1NTTCBDb3Jwb3JhdGlvbjExMC8G
# A1UEAwwoU1NMLmNvbSBSb290IENlcnRpZmljYXRpb24gQXV0aG9yaXR5IFJTQTCC
# AiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAPkP3aMrfcvQKv7sZ4Wm5y4b
# unfh4/WvpOz6Sl2RxFdHaxh3a3by/ZPkPQ/CFp4LZsNWlJ4Xg4XOVu/yFv0AYvUi
# CVToZRdOQbngT0aXqhvIuG5iXmmxX9sqAn78bMrzQdjt0Oj8P2FI7bADFB0QDksZ
# 4LtO7IZl/zbzXmcCC52GVWH9ejjt/uIZALdvoVBidXQ8oPrIJZK0bnoix/geoeOy
# 3ZExqysdBP+lSgQ36YWkMyv94tZVNHwZpEpox7Ko07fKoZOI68GXvIz5HdkihCR0
# xwQ9aqkpk8zruFvh/l8lqjRYyMEjVJ0bmBHDOJx+PYZspQ9AhnwC9FwCTyjLrnGf
# DzrIM/4RJTXq/LrFYD3ZfBjVsqnTdXgDciLKOsMf7yzlLqn6niy2UUb9rwPW6mBo
# 6oUWNmuF6R7As93EJNyAKoFBbZQ+yODJgUEAnl6/f8UImKIYLEJAs/lvOCdLToD0
# PYFH4Ih86hzOtXVcUS4cK38acijnALXRdMbX5J+tB5O2UzU1/Dfkw/ZdFr4hc96S
# CvigY2q8lpJqPvi8ZVWb3vUNiSYE/CUapiVpy8JtynziWV+XrOvvLsi81xtZPCvM
# 8hnIk2snYxnP/Okm+Mpxm3+T/jRnhE6Z6/yzeAkzcLpmpnbtG3PrGqUNxCITIJRW
# Ck4sbE6x/c+cCbqiM+2HAgMBAAGjYzBhMB0GA1UdDgQWBBTdBAkHovV6fVJTEpKV
# 7jiAJQ2mWTAPBgNVHRMBAf8EBTADAQH/MB8GA1UdIwQYMBaAFN0ECQei9Xp9UlMS
# kpXuOIAlDaZZMA4GA1UdDwEB/wQEAwIBhjANBgkqhkiG9w0BAQsFAAOCAgEAIBgR
# lCn7Jp0cHh5wYfGVcpNxJK1ok1iOMq8bs3AD/CUrdIWQPXhq9LmLpZc7tRiRux6n
# +UBbkflVma8eEdBcHadm47GUBwwyOabqG7B52B2ccETjit3E+ZUfijhDPwGFpUen
# PUayvOUiaPd7nNgsPgohyC0zrL/FgZkxdMF1ccW+sfAjRfSda/wZY52jvATGGAsl
# u1OJD7OAUN5F7kR/q5R4ZJjT9ijdh9hwZXT7DrkT66cPYakylszeu+1jTBi7qUD3
# oFRuIIhxdRjqerQ0cuAjJ3dctpDqhiVAq+8zD8ufgr6iIPv2tS0a5sKFsXQP+8hl
# AqRSAUfdSSLBv9jra6x+3uxjMxW3IwiPxg+NQVrdjsW5j+VFP3jbutIbQLH+cU0/
# 4IGiul607BXgk90IH37hVZkLId6Tngr75qNJvTYw/ud3sqB1l7UtgYgXZSD32pAA
# n8lSzDLKNXz1PQ/YK9f1JmzJBjSWFupwWRoyeXkLtoh/D1JIPb9s2KJELtFOt3JY
# 04kTlf5Eq/jXixtunLwsoFvVagCvXzfh1foQC5ichucmj87w7G6KVwuA406ywKBj
# YZC6VWg3dGq2ktufoYYitmUnDuy2n0Jg5GfCtdpBC8TTi2EbvPofkSvXRAdeuims
# 2cXp71NIWuuA8ShYIc2wBlX7Jz9TkHCpBB5XJ7kwggbLMIIEs6ADAgECAhBSnbvC
# uImH9YOA4LDrogY2MA0GCSqGSIb3DQEBCwUAMHwxCzAJBgNVBAYTAlVTMQ4wDAYD
# VQQIDAVUZXhhczEQMA4GA1UEBwwHSG91c3RvbjEYMBYGA1UECgwPU1NMIENvcnBv
# cmF0aW9uMTEwLwYDVQQDDChTU0wuY29tIFJvb3QgQ2VydGlmaWNhdGlvbiBBdXRo
# b3JpdHkgUlNBMB4XDTI2MDYxMjE2NDIxNVoXDTQxMDIxMjE3MzkzOVowcjELMAkG
# A1UEBhMCVVMxDjAMBgNVBAgMBVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMREwDwYD
# VQQKDAhTU0wgQ29ycDEuMCwGA1UEAwwlU1NMLmNvbSBUaW1lc3RhbXBpbmcgSXNz
# dWluZyBSU0EgQ0EgMjCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAJ8v
# q+ZzTZC8MUabNgd5WlBjVPcu71oiedNd8n9Il8+uW0HcGvW44/Jf0pHl6hxqn0od
# haGnvvNBPJgxH/d6CMI7Mtgl9nF5isarRfByRmAeH+skYShA+dMJGzsrC9NiF3/v
# xGZ2uOI334YVc+DlP5OdsUPz5yIP7hYXNlpPknPwh6gK/pWFZD8+y8uUagfX9USh
# q/SQt+1tS4FXRwudHTSHbGt0Fz4ZbWGbm0d56ROAW34tI6FFuZ+TtGssHjeOUf2F
# S8QNGBpadTh1X1Z7Kl6vG4Z8NS3OJIyelzoF5TpcP2vM3ue0723bmZd3oXsE5rzt
# AJMv56jlg0vSDz8XcXRur6awW0rh5cTQDQQGx04hYJH2Zn9HWWJU/wJ2SKBcQKtB
# yCz6JX+DoFoBE++kolkpZtl2ESF8FY6XqhtKZb18ESEAc0ohCY+0HkBZwrx8eIED
# CDbzd/shv9a8YRG2oxr3tLUz95V+Dzzh1QBNECaeGczyNWGCZIVUa4lVcVUVOBbI
# OAT8D+iPmC5h29bWeES8DhTS4f4CWA4D5QoQqMTqIBcV4MD2Waw8hEWAdGnjnsVt
# DimnDeEUbvXKhdiirvwG+QC+3mAy1RlD6xBMmGyDAKVSqqXm3NP4w0J4gvt65Grk
# QI4FXPJtYqfwkIBtQ8OGfbpVUQGPVUemWtui7E21AgMBAAGjggFRMIIBTTASBgNV
# HRMBAf8ECDAGAQH/AgEAMB8GA1UdIwQYMBaAFN0ECQei9Xp9UlMSkpXuOIAlDaZZ
# ME0GCCsGAQUFBwEBBEEwPzA9BggrBgEFBQcwAoYxaHR0cDovL2NlcnQuc3NsLmNv
# bS9TU0xjb20tUm9vdENBLVJTQS00MDk2LVIxLmNlcjA/BgNVHSAEODA2MDQGBFUd
# IAAwLDAqBggrBgEFBQcCARYeaHR0cHM6Ly93d3cuc3NsLmNvbS9yZXBvc2l0b3J5
# MBMGA1UdJQQMMAoGCCsGAQUFBwMIMEIGA1UdHwQ7MDkwN6A1oDOGMWh0dHA6Ly9j
# cmxzLnNzbC5jb20vU1NMY29tLVJvb3RDQS1SU0EtNDA5Ni1SMS5jcmwwHQYDVR0O
# BBYEFHCaZcFDgKP86bFkiW7xMJwJIcl8MA4GA1UdDwEB/wQEAwIBhjANBgkqhkiG
# 9w0BAQsFAAOCAgEAZRRbzTHpgOogYgWKOeiCkI1oVrkGmCrqOKqX9CiTREQLdYM/
# h7ANki98QP1fhAPEkTzT/Ur19IrPuZg7KaWL6SsmBqmD9+mBVtJmxdRdGbV8bpgl
# FDv2VeZgTSg3+UOD2O0G9HlgHqVHNKzSa/zL5+WYl3yT8Q29sZnIPGkQ+udv56/j
# q5NlAT9BON4iXSOCxe1bNvUd0hkS2GTsTNeQyOEHm7CkZdCg5X+DpmBginO7pk8Y
# m5GbTdQtSloJ3/Y2nVPycJ8kmmBAYMev7/hVQukhjjc4mO03LyPTFfqmMR1D5Tpu
# gGkD9nUuG41l55vOpXQm1JMbTgArpkn0/e9O+yTYlWuz2dSAZHf9jxNQAvILO1Pv
# 2JNQTfV0TyIteZP7gNcRfaHnnuSho8ohipKiE+1vEJIFE+c3gopxSpHCmn2THETk
# lWKV7ZrUKhhkdoOLj2KidCtYZrOoCV6x7f9+KfDPT61fKDVntGZdRKwzLV53VIN6
# oSS+KQN2VJaO8OZTOZpfgO4o6/9xpNr+EEaA3hNkoZk13R4LaqPC/309dfZmHBE8
# nBNjWBkV7AZdzdK8w7B0a/D9066a6oUa19OIktD+3kPpQ8YEt8oRp9zs1hoQ7fRn
# 9sdantP6nYRQFTpF8czGc+FGresbAH819Pv5vBhCluTo8Px2PpC71cM0jzsxggJV
# MIICUQIBATCBhjByMQswCQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNV
# BAcMB0hvdXN0b24xETAPBgNVBAoMCFNTTCBDb3JwMS4wLAYDVQQDDCVTU0wuY29t
# IFRpbWVzdGFtcGluZyBJc3N1aW5nIFJTQSBDQSAyAhBUQ/g3vVruGAGEcrhc2o7j
# MAsGCWCGSAFlAwQCAaCCAWAwGgYJKoZIhvcNAQkDMQ0GCyqGSIb3DQEJEAEEMBwG
# CSqGSIb3DQEJBTEPFw0yNjExMDUxNzUxMjJaMCgGCSqGSIb3DQEJNDEbMBkwCwYJ
# YIZIAWUDBAIBoQoGCCqGSM49BAMCMC8GCSqGSIb3DQEJBDEiBCAOnNs0I3ra8NcT
# Q9XjYVUpe+ZwInZN6Pf7KES+tFDONDCByAYLKoZIhvcNAQkQAi8xgbgwgbUwgbIw
# ga8EIFdP6heq+l58aiubCMmYsXNeRQs4c8pE5qJMoJMOY/OwMIGKMHakdDByMQsw
# CQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNVBAcMB0hvdXN0b24xETAP
# BgNVBAoMCFNTTCBDb3JwMS4wLAYDVQQDDCVTU0wuY29tIFRpbWVzdGFtcGluZyBJ
# c3N1aW5nIFJTQSBDQSAyAhBUQ/g3vVruGAGEcrhc2o7jMAoGCCqGSM49BAMCBEYw
# RAIgOox0gYckCxG6iSqDZ8olh9jQg3mweq5jkyz3NxEh7ywCIBpvE5/ZOQF6z+zL
# lMKkm2wt8dx3V7NH9Yd373KtQF1U
# SIG # End signature block
