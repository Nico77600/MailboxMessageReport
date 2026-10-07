<#
.SYNOPSIS
    Renders the images of the guides and the readme: the window (light and dark, after a search, during a search) and the
    HTML report (light and dark, the detail of a message, the summary of the per-mailbox layout), from fictitious data.

.DESCRIPTION
    No tenant and no real data: the mailboxes come from the simulated tenant of the tests
    (tests\MailboxMessageReport.FakeGraph.ps1), loaded inside the module, with contoso.com names. The window is rendered
    off screen (RenderTargetBitmap); the report is opened by Microsoft Edge headless.

    Writes docs\images\gui-search-light.png, gui-search-dark.png, gui-progress-light.png, report-overview.png,
    report-dark.png, report-message.png, report-mailboxes.png, gui-list-light.png. Needs an interactive session (WPF) and Microsoft Edge.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.1.0
#>
#Requires -Version 7.4
[CmdletBinding()]
param([string]$Destination = (Join-Path $PSScriptRoot '..\docs\images'))

$ErrorActionPreference = 'Stop'
# The images of the English documentation: the texts of WPF (the empty date field) in English.
[cultureinfo]::CurrentUICulture = [cultureinfo]::CurrentCulture = [cultureinfo]'en-US'
$root = Split-Path $PSScriptRoot -Parent
$Destination = [IO.Path]::GetFullPath($Destination)
[void][IO.Directory]::CreateDirectory($Destination)
# The reports of the images go where a real installation would put them (the path shows in the window), only when that
# folder does not exist (it is removed at the end); otherwise under artifacts\.
$neutral = Join-Path $env:SystemDrive 'Tools\MailboxMessageReport'
$ownNeutral = -not (Test-Path -LiteralPath $neutral)
$ownParent = -not (Test-Path -LiteralPath (Split-Path $neutral -Parent))
$work = if ($ownNeutral) { Join-Path $neutral 'reports' } else { Join-Path $root 'artifacts\doc-images' }
if (-not $ownNeutral -and (Test-Path $work)) { Remove-Item $work -Recurse -Force }
try { [void][IO.Directory]::CreateDirectory($work) }
catch { $ownNeutral = $false; $work = Join-Path $root 'artifacts\doc-images'; if (Test-Path $work) { Remove-Item $work -Recurse -Force }; [void][IO.Directory]::CreateDirectory($work) }
$cleanup = {
    if ($ownNeutral) {
        Start-Sleep -Seconds 1
        Remove-Item -LiteralPath ($(if ($ownParent) { Split-Path $neutral -Parent } else { $neutral })) -Recurse -Force -ErrorAction SilentlyContinue
    }
}
trap { & $cleanup; break }
Import-Module (Join-Path $root 'MailboxMessageReport.psd1') -Force
$module = Get-Module MailboxMessageReport

& $module {
    param($FakePath, $Destination, $Work)
    Set-StrictMode -Off
    . $FakePath
    function script:Start-MmrGraphSend { param($Method, $Url, $Body, $Headers) [pscustomobject]@{ Task = $null; Request = $null; Response = (Invoke-FakeGraphHttp -Method $Method -Url $Url -Body $Body -Headers $Headers) } }
    $script:Quiet = $true
    # The simulated tenant lives in this runspace: the window runs its work here (not in its background runspace).
    $script:GuiInline = $true

    # ---- fictitious tenant: Megan and Alex have an archive, Lee Gu does not ---------------------------------------
    $tenant = '0b6c7d1e-2f3a-4b5c-8d9e-0f1a2b3c4d5e'
    Reset-FakeTenant -TenantId $tenant
    $d = 'contoso.com'
    $null = Add-FakeUser "megan.bowen@$d" -Name 'Megan Bowen' -Aliases "mbowen@$d" -Archive
    $null = Add-FakeUser "alex.wilber@$d" -Name 'Alex Wilber' -Archive
    $null = Add-FakeUser "lee.gu@$d" -Name 'Lee Gu'
    $rnd = [Random]::new(2026)
    $people = @("adele.vance@$d", "lidia.holloway@$d", "nestor.wilke@$d", "joni.sherman@$d", 'jacques.martin@fabrikam.com', 'claire.dupont@northwind.com', 'pradeep.gupta@tailspin.com')
    $subjects = @('Contrat Alpha - version {0}', 'Projet Atlas : compte rendu n°{0}', 'Facture {1}-{0:000}', 'Budget {1} - arbitrage', 'RE: Offre commerciale Fabrikam ({0})', 'Weekly status #{0}', 'Planning des congés {1}', 'Revue fournisseur Northwind', 'TR: Contrat Alpha - signature', 'Kick-off Atlas : présentation', 'Accord de confidentialité Tailspin', 'Note de frais {1}/{0:00}')
    $add = {
        param([string]$User, [string]$Location, [string]$Path, [int]$Count, [datetime]$From, [datetime]$To, [switch]$Recoverable, [switch]$Sent)
        $span = ($To - $From).TotalMinutes
        for ($i = 1; $i -le $Count; $i++) {
            $date = $From.AddMinutes($rnd.NextDouble() * $span)
            $other = $people[$rnd.Next($people.Count)]
            $toList = if ($Sent) { @($other) + @($people[$rnd.Next($people.Count)]) } else { @($User) + $(if ($rnd.Next(3) -eq 0) { @($people[$rnd.Next($people.Count)]) } else { @() }) }
            $cc = if ($rnd.Next(3) -eq 0) { @($people[$rnd.Next($people.Count)]) } else { @() }
            $bcc = if ($Sent -and $rnd.Next(5) -eq 0) { @("assistant@$d") } else { @() }
            # Not $from or $to: PowerShell variables ignore the case, $From and $To are the period.
            $sender = if ($Sent) { $User } else { $other }
            $name = (Get-Culture).TextInfo.ToTitleCase(($sender -split '@')[0].Replace('.', ' '))
            $null = Add-FakeMessage $User -Location $Location -Path $Path -Recoverable:$Recoverable -Subject ($subjects[$rnd.Next($subjects.Count)] -f $rnd.Next(1, 60), $date.Year) -Received $date -From $sender -FromName $name -To $toList -Cc $cc -Bcc $bcc -Attachments:($rnd.Next(4) -eq 0)
        }
    }
    foreach ($pair in @(@("megan.bowen@$d", 1.0), @("alex.wilber@$d", 0.6))) {
        $u = $pair[0]; $k = $pair[1]
        & $add $u Primary '\Inbox' ([int](24 * $k)) ([datetime]'2026-06-01') ([datetime]'2026-10-06')
        & $add $u Primary '\Inbox\Projects\Atlas' ([int](9 * $k)) ([datetime]'2026-02-01') ([datetime]'2026-09-30')
        & $add $u Primary '\Sent Items' ([int](12 * $k)) ([datetime]'2026-05-01') ([datetime]'2026-10-06') -Sent
        & $add $u Archive '\Inbox' ([int](30 / $k)) ([datetime]'2019-01-01') ([datetime]'2024-12-31')
        & $add $u Archive '\Inbox\Contrats' 14 ([datetime]'2020-03-01') ([datetime]'2024-11-30')
        & $add $u Archive '\Sent Items' ([int](16 * $k)) ([datetime]'2019-01-01') ([datetime]'2024-12-31') -Sent
        & $add $u Archive '\2018' ([int](18 / $k)) ([datetime]'2018-01-02') ([datetime]'2018-12-28')
        & $add $u Primary '\Deletions' ([int](4 * $k) + 1) ([datetime]'2025-09-01') ([datetime]'2026-08-30') -Recoverable
        $null = Add-FakeFolder $u -Path '\Purges' -Recoverable
    }
    $atlas = '<p>Bonjour Megan,</p><p>Voici le compte rendu du comité Atlas de ce matin.</p><p>1. Planning : la bascule est confirmée pour le 14 novembre, la répétition générale le 7.</p><p>2. Budget : le dépassement de 4 % est validé par la direction financière.</p><p>3. Risques : la migration des archives est le point d''attention ; Alex prépare le rapport des boîtes et de leurs archives d''ici vendredi.</p><p>Prochain comité le 21 octobre.</p><p>Bonne journée,<br>Lidia</p>'
    $null = Add-FakeMessage "megan.bowen@$d" -Path '\Inbox' -Subject 'Projet Atlas : compte rendu du comité' -Received ([datetime]'2026-10-06T16:42:00') -From "lidia.holloway@$d" -FromName 'Lidia Holloway' -To "megan.bowen@$d", "alex.wilber@$d" -Cc "adele.vance@$d" -Attachments -Body $atlas
    & $add "lee.gu@$d" Primary '\Inbox' 20 ([datetime]'2025-11-01') ([datetime]'2026-10-06')
    & $add "lee.gu@$d" Primary '\Sent Items' 8 ([datetime]'2025-11-01') ([datetime]'2026-10-06') -Sent
    $null = Add-FakeFolder "lee.gu@$d" -Path '\Junk Email'

    $settings = Get-MmrDefaultConfiguration
    $settings.TenantId = 'contoso.onmicrosoft.com'; $settings.Organization = 'contoso.onmicrosoft.com'; $settings.AppId = '6b8e1f0a-3c52-4b8e-9a51-0f2d7c3e4a19'
    $settings.CertificateThumbprint = '3F2A9C7B1E6D4A8F0B5C2E9D7A1F4B6C8E0D2A5B'; $settings.TimeZone = 'Europe/Paris'
    $settings.OutputPath = $Work; $settings.LogPath = Join-Path $Work 'logs'; $settings.ConfigPath = 'config\MailboxMessageReport.config.psd1'
    $token = New-FakeToken -TenantId $tenant
    $script:Graph = @{ Settings = $settings; Token = $token; ExpiresUtc = [datetime]::UtcNow.AddHours(1); Certificate = $null; Secret = $null; Roles = @('Mail.Read', 'User.Read.All')
        CanReadMail = $true; CanReadUsers = $true; CanFindUsers = $true; TenantGuid = $tenant; AppName = 'Mailbox Message Report'; Renew = { @{ Token = $token; ExpiresUtc = [datetime]::UtcNow.AddHours(1) } } }
    function Connect-MmrGraph { param($Settings, $Secret) [pscustomobject]$script:Graph }

    $render = {
        param($Form, [string]$Path)
        $w = $Form.Form
        $w.WindowStartupLocation = 'Manual'; $w.Left = -4000; $w.Top = 0; $w.Width = 1480; $w.Height = 940; $w.ShowInTaskbar = $false
        if (-not $w.IsVisible) { $w.Show() }
        [Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([Action] {}, [Windows.Threading.DispatcherPriority]::Background)
        $w.UpdateLayout()
        Update-MmrGuiColumns
        [Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([Action] {}, [Windows.Threading.DispatcherPriority]::Background)
        $w.UpdateLayout()
        $rtb = [Windows.Media.Imaging.RenderTargetBitmap]::new([int]$w.ActualWidth, [int]$w.ActualHeight, 96, 96, [Windows.Media.PixelFormats]::Pbgra32)
        $rtb.Render($w.Content)
        $enc = [Windows.Media.Imaging.PngBitmapEncoder]::new(); $enc.Frames.Add([Windows.Media.Imaging.BitmapFrame]::Create($rtb))
        $fs = [IO.File]::Create($Path); try { $enc.Save($fs) } finally { $fs.Dispose() }
    }
    $search = {
        param([string]$Theme, [string]$Layout = 'Global')
        $f = New-MmrForm -Configuration $settings -Theme $Theme
        $f.Form.WindowStartupLocation = 'Manual'; $f.Form.Left = -4000; $f.Form.ShowInTaskbar = $false; $f.Form.Show()
        $f.Controls.Mailbox.Text = "megan.bowen@$d" + [Environment]::NewLine + "alex.wilber@$d" + [Environment]::NewLine + "lee.gu@$d"
        $f.Controls.RecoverableItems.IsChecked = $true
        if ($Layout -eq 'PerMailbox') { $f.Controls.LayoutPerMailbox.IsChecked = $true }
        $f.Controls.ConnectionExpander.IsExpanded = $false
        Invoke-MmrGuiSearch
        # The content of the message selected in the reading pane (read at once: no timer off screen).
        Read-MmrGuiBody
        $f
    }

    $light = & $search 'Light'
    & $render $light (Join-Path $Destination 'gui-search-light.png')
    $reportFolder = $script:Gui.LastFolder
    # The list view of the same preview.
    $light.Controls.ViewList.IsChecked = $true
    & $render $light (Join-Path $Destination 'gui-list-light.png')
    $light.Form.Close()
    $dark = & $search 'Dark'
    & $render $dark (Join-Path $Destination 'gui-search-dark.png')
    $dark.Form.Close()

    # ---- a run in progress: the lines of a run on a large tenant ---------------------------------------------------
    $running = New-MmrForm -Configuration $settings -Theme Light
    $running.Form.WindowStartupLocation = 'Manual'; $running.Form.Left = -4000; $running.Form.ShowInTaskbar = $false; $running.Form.Show()
    $running.Controls.Mailbox.Text = "megan.bowen@$d" + [Environment]::NewLine + "alex.wilber@$d" + [Environment]::NewLine + "lee.gu@$d"
    $running.Controls.StartDate.SelectedDate = [datetime]'2019-01-01'
    $running.Controls.Subject.Text = 'Contrat Alpha'
    $running.Controls.ConnectionExpander.IsExpanded = $false
    Start-MmrGuiRun 'Reading...'
    foreach ($line in @(
            @('Step', '[1/5] Microsoft Graph'), @('Ok', 'Application Mailbox Message Report · tenant contoso.onmicrosoft.com'), @('Info', 'Permissions: Mail.ReadBasic.All, User.Read.All'),
            @('Step', '[2/5] Mailboxes (3)'), @('Ok', "Megan Bowen <megan.bowen@$d> · primary mailbox · archive"), @('Ok', "Alex Wilber <alex.wilber@$d> · primary mailbox · archive"), @('Ok', "Lee Gu <lee.gu@$d> · primary mailbox · no archive"),
            @('Step', '[3/5] Folders'), @('Ok', 'Primary mailbox: 46 folders · 31 to read · 15 empty'), @('Ok', 'Archive: 38 folders · 29 to read · 9 empty'),
            @('Step', '[4/5] Messages'), @('Info', "received from 2019-01-01 00:00 · subject contains 'Contrat Alpha'"), @('Info', '2 large folder(s) read in 14 slices of their dates, side by side'),
            @('Progress', '0.612|38/60 folders read · 1,284 messages|about 1 min 10 s left'))) { Add-MmrGuiLine $line[0] $line[1] }
    & $render $running (Join-Path $Destination 'gui-progress-light.png')
    Stop-MmrGuiRun
    $running.Form.Close()

    # ---- one report per mailbox: the summary and its links ---------------------------------------------------------
    $per = & $search 'Light' 'PerMailbox'
    $perFolder = $script:Gui.LastFolder
    $per.Form.Close()
    [pscustomobject]@{ Report = (Join-Path $reportFolder 'MailboxMessageReport.html'); PerMailbox = (Join-Path $perFolder 'MailboxMessageReport.html') }
} (Join-Path $root 'tests\MailboxMessageReport.FakeGraph.ps1') $Destination $work | Set-Variable reports

# ---- the HTML report, opened by Microsoft Edge headless ------------------------------------------------
$edge = @("${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe", "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe") | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $edge) { throw 'Microsoft Edge is needed for the images of the report.' }
foreach ($shot in @(
        @{ Html = $reports.Report; Theme = 'light'; File = 'report-overview.png' }
        @{ Html = $reports.Report; Theme = 'dark'; File = 'report-dark.png' }
        @{ Html = $reports.Report; Theme = 'light'; File = 'report-message.png'; Query = '&scoutFocus=tables&scoutDetail=4'; Size = '1360,900' }
        @{ Html = $reports.PerMailbox; Theme = 'light'; File = 'report-mailboxes.png'; Query = '&scoutFocus=tables'; Size = '1360,520' })) {
    $profile = Join-Path $work "edge-$([guid]::NewGuid().ToString('N'))"
    $url = ([Uri]$shot.Html).AbsoluteUri + "?scoutTheme=$($shot.Theme)" + $(if ($shot.Query) { $shot.Query } else { '' })
    $size = if ($shot.Size) { $shot.Size } else { '1360,1180' }
    $png = Join-Path $Destination $shot.File
    $before = if (Test-Path -LiteralPath $png) { (Get-Item -LiteralPath $png).LastWriteTimeUtc } else { [datetime]::MinValue }
    # Edge headless sometimes stays open after writing its screenshot: waited for 60 s at most, then stopped.
    $p = Start-Process -FilePath $edge -PassThru -WindowStyle Hidden -ArgumentList @('--headless=new', '--disable-gpu', '--hide-scrollbars', "--user-data-dir=`"$profile`"", "--window-size=$size", '--virtual-time-budget=3000', "--screenshot=`"$png`"", "`"$url`"")
    if (-not $p.WaitForExit(60000)) {
        if (-not (Test-Path -LiteralPath $png) -or (Get-Item -LiteralPath $png).LastWriteTimeUtc -le $before) { Write-Warning "Edge did not write $($shot.File) within 60 s." }
        Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue
    }
    Start-Sleep -Milliseconds 500
}
& $cleanup
Get-ChildItem -LiteralPath $Destination -Filter '*.png' | Where-Object Name -notlike 'readme-*' | Sort-Object Name | ForEach-Object { '{0,10:N0}  {1}' -f $_.Length, $_.Name }
