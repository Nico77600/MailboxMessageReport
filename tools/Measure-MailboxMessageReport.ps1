<#
.SYNOPSIS
    Measures the time of each step of Mailbox Message Report once Microsoft Graph has answered, on a large synthetic
    volume, without a tenant.

.DESCRIPTION
    The time to read the messages depends on Exchange Online (a page of 1,000 messages takes about 1 to 3 s, 4 pages at
    a time per mailbox): see the developer guide, chapter 13. This tool measures what the computer adds to it:

      1. Pages      pages of 1,000 synthetic messages as Graph answers them (every column of the report filled) written
                    to the part files of the folders (MailboxMessageReportNative.PartWriter): the work done for each
                    page while the next ones are on their way;
      2. Report     the part files merged into the CSV file, the first rows of the HTML report and the preview of the
                    window, the HTML report written (Export-MmrReport, Global layout);
      3. Memory     the working set of the process before and after: the messages are on disk, not in memory.
    With -Simulated, the same search end to end against the simulated tenant of the tests (smaller volume: the
    simulated tenant is written in PowerShell and is much slower than Graph): the cost of the scheduler.

.PARAMETER Messages
    Messages written (default 200,000), in folders of -FolderSize messages, in -Mailboxes mailboxes.

.PARAMETER HtmlMaxMessages
    Messages of the HTML report (Report.HtmlMaxMessages; default 20,000).

.PARAMETER Simulated
    Also a search of the simulated tenant: -SimulatedMailboxes mailboxes with an archive, 5 folders each side, 200
    messages per folder.

.EXAMPLE
    pwsh -File .\tools\Measure-MailboxMessageReport.ps1 -Messages 500000 -Simulated

.NOTES
    Author  : Nicolas Fabert
    Version : 1.1.1
#>
#Requires -Version 7.4
[CmdletBinding()]
param(
    [int]$Messages = 200000,
    [int]$FolderSize = 5000,
    [int]$Mailboxes = 10,
    [int]$HtmlMaxMessages = 20000,
    [switch]$Simulated,
    [int]$SimulatedMailboxes = 10
)

$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
Import-Module (Join-Path $root 'MailboxMessageReport.psd1') -Force
$module = Get-Module MailboxMessageReport
$work = Join-Path $root 'artifacts\measure'
if (Test-Path $work) { Remove-Item $work -Recurse -Force }
[void][IO.Directory]::CreateDirectory($work)
$inv = [Globalization.CultureInfo]::InvariantCulture
$results = [Collections.Generic.List[object]]::new()
$add = { param([string]$Step, [double]$Seconds, [long]$Count, [string]$Note) $results.Add([pscustomobject]@{ Step = $Step; Seconds = [Math]::Round($Seconds, 2); PerSecond = $(if ($Seconds -gt 0 -and $Count) { [long]($Count / $Seconds) } else { 0 }); Note = $Note }) }
$memory = { [Math]::Round([Diagnostics.Process]::GetCurrentProcess().WorkingSet64 / 1MB) }

# ---- a page of 1,000 messages, as Graph answers it --------------------------------------------------------------
$people = 'alice.martin', 'bob.durand', 'chloe.bernard', 'david.petit', 'emma.roux', 'franck.morel', 'gabriel.simon'
$address = { param([int]$i) @{ emailAddress = @{ name = (Get-Culture).TextInfo.ToTitleCase($people[$i % $people.Count].Replace('.', ' ')); address = "$($people[$i % $people.Count])@contoso.com" } } }
$value = for ($i = 0; $i -lt 1000; $i++) {
    $d = ([datetime]'2020-01-01').AddMinutes(-$i * 37)
    [ordered]@{
        '@odata.etag' = 'W/"CQAAABYAAAAGWiHnwH0mRblaqPsoBtheAABWFy0f"'
        id = 'AAMkADAxNjM3MWJjLWZmYTMtNGJhYy1iNmIzLWQxMTBhMmYyMjYzNQBGAAAAAABX3YUBda1YQqgX_-c0zooIBwAGWiHnwH0mRblaqPsoBtheAABWF0eHAAAGWiHnwH0mRblaqPsoBtheAABWF' + $i.ToString('D5', $inv)
        receivedDateTime = $d.ToString('yyyy-MM-ddTHH:mm:ssZ', $inv); sentDateTime = $d.AddMinutes(-1).ToString('yyyy-MM-ddTHH:mm:ssZ', $inv)
        hasAttachments = ($i % 5 -eq 0); internetMessageId = "<mmr-measure-$i-$([guid]::NewGuid().ToString('n'))@contoso.com>"; subject = "Contrat Alpha - version $i (review of the budget $($d.Year))"
        importance = 'normal'; isRead = $true
        sender = (& $address $i); from = (& $address $i)
        toRecipients = @((& $address ($i + 1)), (& $address ($i + 2))); ccRecipients = @((& $address ($i + 3))); bccRecipients = @()
    }
}
$page = @{ '@odata.context' = 'https://graph.microsoft.com/v1.0/$metadata#users(...)/mailFolders(...)/messages(...)'; value = @($value); '@odata.nextLink' = 'https://graph.microsoft.com/v1.0/users/x/mailFolders/y/messages?$skip=1000' } | ConvertTo-Json -Depth 8 -Compress
# The page as the transport gives it: its bytes (a Body), as read from the answer of Graph.
$body = [MailboxMessageReportNative.Body]::new([Text.Encoding]::UTF8.GetBytes($page))
$pages = [int][Math]::Ceiling($Messages / 1000)
$perFolder = [Math]::Max(1, [int][Math]::Ceiling($FolderSize / 1000))
Write-Host ''
Write-Host ("  Mailbox Message Report {0} - measure: {1:N0} messages, folders of {2:N0}, {3} mailboxes, page {4:N0} KB" -f $module.Version, ($pages * 1000), ($perFolder * 1000), $Mailboxes, [Math]::Round($page.Length / 1KB)) -ForegroundColor Cyan
$before = & $memory

# ---- 1. pages written to the part files ------------------------------------------------------------------------
$parts = Join-Path $work '.parts'
[void][IO.Directory]::CreateDirectory($parts)
$zone = [TimeZoneInfo]::FindSystemTimeZoneById('Romance Standard Time')
$folders = [Collections.Generic.List[object]]::new()
$sw = [Diagnostics.Stopwatch]::StartNew()
$written = [long]0; $n = 0
for ($p = 0; $p -lt $pages; $p += $perFolder) {
    $mailbox = "user$('{0:D3}' -f ($n % $Mailboxes))@contoso.com"
    $location = if ($n % 3 -eq 0) { 'Primary' } else { 'Archive' }
    $path = '\Archive {0}\Folder {1:D4}' -f (2010 + $n % 12), $n
    $file = Join-Path $parts ('{0:D6}.jsonl' -f $n)
    $w = [MailboxMessageReportNative.PartWriter]::new($file, $mailbox, "User $($n % $Mailboxes)", $location, $false, $path, "Folder $n", $zone)
    try { for ($k = 0; $k -lt $perFolder -and $p + $k -lt $pages; $k++) { $written += $w.AddPage($body).Rows } }
    finally { $w.Dispose() }
    $folders.Add([pscustomobject]@{ Mailbox = $mailbox; MailboxName = "User $($n % $Mailboxes)"; MailboxKey = $mailbox; Location = $location; RecoverableItems = $false; Id = "f$n"; Name = "Folder $n"; Path = $path; TotalItems = [long]($perFolder * 1000); Messages = [long]$w.Count; Status = 'Read'; Detail = ''; Parts = @($file); Slices = 1; SlicesDone = 1; Pages = $perFolder })
    $n++
}
$sw.Stop()
$partBytes = (Get-ChildItem $parts -File | Measure-Object Length -Sum).Sum
& $add 'Pages to part files' $sw.Elapsed.TotalSeconds $written ('{0:N0} pages, {1:N0} folders, {2:N0} MB of part files' -f $pages, $folders.Count, [Math]::Round($partBytes / 1MB))

# ---- 2. the report ---------------------------------------------------------------------------------------------------
$result = & $module {
    param($folders, $Mailboxes)
    $s = Get-MmrDefaultConfiguration
    $s.TimeZone = 'Romance Standard Time'
    $request = [pscustomobject]@{ Mailboxes = @(); MailboxFile = ''; Start = $null; End = $null; Subject = @(); Location = @('Primary', 'Archive'); RecoverableItems = $false; ExcludeFolders = @(); Layout = 'Global'; Formats = @('Csv', 'Html'); TimeZone = 'Romance Standard Time' }
    $r = New-MmrResult -Settings $s -Request $request
    $r.Mailboxes = @(0..($Mailboxes - 1) | ForEach-Object { $m = New-MmrMailbox -Address ('user{0:D3}@contoso.com' -f $_); $m.DisplayName = "User $_"; $m.Archive = 'Yes'; $m })
    $r.Folders = @($folders)
    foreach ($m in $r.Mailboxes) {
        foreach ($f in @($folders | Where-Object Mailbox -eq $m.Address)) { $m.FoldersRead++; if ($f.Location -eq 'Archive') { $m.ArchiveMessages += $f.Messages } else { $m.PrimaryMessages += $f.Messages }; $m.Messages += $f.Messages }
    }
    Update-MmrResultCounts $r
    $r
} $folders $Mailboxes
& $module { $script:Quiet = $true }
$sw = [Diagnostics.Stopwatch]::StartNew()
$report = Export-MmrReport -Result $result -Directory $work -Formats Csv, Html -Layout Global -HtmlMaxMessages $HtmlMaxMessages -PreviewMessages 1000 -PartsPath $parts
$sw.Stop()
& $add 'Report (CSV, HTML, preview)' $sw.Elapsed.TotalSeconds $written ('Messages.csv {0:N0} MB, HTML {1:N0} MB (first {2:N0})' -f [Math]::Round((Get-Item $report.Files.Messages).Length / 1MB), [Math]::Round((Get-Item $report.Files.Html).Length / 1MB, 1), $HtmlMaxMessages)
$sw = [Diagnostics.Stopwatch]::StartNew()
$rows = [MailboxMessageReportNative.PreviewRow]::Build($report.Preview)
$sw.Stop()
& $add 'Rows of the window' $sw.Elapsed.TotalSeconds $rows.Count ('{0:N0} rows of preview' -f $rows.Count)
$after = & $memory

# ---- 3. a search of the simulated tenant -------------------------------------------------------------------------
if ($Simulated) {
    . (Join-Path $root 'tests\MailboxMessageReport.FakeGraph.ps1')
    Reset-FakeTenant
    $total = 0
    for ($i = 0; $i -lt $SimulatedMailboxes; $i++) {
        $u = 'sim{0:D3}@contoso.test' -f $i
        $null = Add-FakeUser $u -Archive
        foreach ($loc in 'Primary', 'Archive') { foreach ($f in 1..5) { foreach ($j in 1..200) { $null = Add-FakeMessage $u -Location $loc -Path "\Folder $f" -Subject "Message $j" -Received ([datetime]'2024-01-01').AddHours(-$j * 7) -To $u; $total++ } } }
    }
    $fake = $script:Fake
    & $module {
        param($fake, $FakePath)
        . $FakePath
        $script:Fake = $fake
        function script:Start-MmrGraphSend { param($Method, $Url, $Body, $Headers) [pscustomobject]@{ Task = $null; Request = $null; Response = (Invoke-FakeGraphHttp -Method $Method -Url $Url -Body $Body -Headers $Headers) } }
        $token = New-FakeToken
        $script:Graph = @{ Settings = (Get-MmrDefaultConfiguration); Token = $token; ExpiresUtc = [datetime]::UtcNow.AddHours(1); Roles = @('Mail.ReadBasic.All', 'User.Read.All'); CanReadMail = $true; CanReadUsers = $true; CanFindUsers = $true; TenantGuid = $fake.Tenant; AppName = 'Measure' }
        $script:Graph.Settings.PageSize = 100
    } $fake (Join-Path $root 'tests\MailboxMessageReport.FakeGraph.ps1')
    $s = & $module { $script:Graph.Settings }
    $request = New-MmrRequest -Settings $s -Mailbox @(0..($SimulatedMailboxes - 1) | ForEach-Object { 'sim{0:D3}@contoso.test' -f $_ })
    $simWork = Join-Path $work 'simulated'
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $sim = & $module { param($s, $r, $p) Initialize-MmrSteps -Total 5; Find-MmrMessages -Settings $s -Request $r -PartsPath $p } $s $request (Join-Path $simWork '.parts')
    $sw.Stop()
    & $add 'Simulated search (scheduler + tenant)' $sw.Elapsed.TotalSeconds $sim.Counts.Messages ('{0} mailboxes and archives, {1:N0} folders, pages of 100' -f $SimulatedMailboxes, $sim.Counts.FoldersRead)
}

Write-Host ''
$results | Format-Table -AutoSize | Out-String -Width 200 | Write-Host
Write-Host ("  Memory (working set): {0:N0} MB before, {1:N0} MB after the report of {2:N0} messages." -f $before, $after, $written)
Write-Host ''
