<#
.SYNOPSIS
    Mailbox Message Report - console output and log file (dot-sourced by MailboxMessageReport.psm1).

.DESCRIPTION
    Same rules as Meeting Cleanup, EAS OAuth Mailbox and Message Trace Report:
      - ANSI colours are disabled when the output is redirected or NO_COLOR is set;
        MMR_FORCE_COLOR=1 forces them.
      - Icons: emoji in Windows Terminal / VS Code, symbols of the classic console fonts elsewhere.
        MMR_ICONS = Emoji | Symbols | Ascii forces a style.
      - Every line shown is also written to the daily log file, without colours or icons.
      - During a window run, the same lines are sent to the progress box of the window.

.NOTES
    Author  : Nicolas Fabert
    Version : 2.1.0
#>

$script:C = @{ Reset = ''; Bold = ''; Dim = ''; Accent = ''; AccentBg = ''; Green = ''; Yellow = ''; Red = ''; Blue = ''; White = '' }
if ($env:Mmr_FORCE_COLOR -eq '1' -or (-not [Console]::IsOutputRedirected -and -not $env:NO_COLOR)) {
    $e = [char]27
    $script:C = @{
        Reset = "$e[0m"; Bold = "$e[1m"; Dim = "$e[90m"; White = "$e[97m"
        Accent = "$e[38;2;214;62;115m"; AccentBg = "$e[48;2;177;31;75m$e[97m"
        Green = "$e[38;2;80;200;120m"; Yellow = "$e[38;2;240;200;90m"; Red = "$e[38;2;240;90;90m"; Blue = "$e[38;2;110;170;240m"
    }
}
$script:IconStyle = if ($env:Mmr_ICONS -in 'Emoji', 'Symbols', 'Ascii') { $env:Mmr_ICONS }
    elseif ([Console]::IsOutputRedirected) { 'Symbols' }
    elseif ($env:WT_SESSION -or $env:TERM_PROGRAM -eq 'vscode') { 'Emoji' }
    else { 'Symbols' }
$script:Dot = [char]0x00B7
$script:ProgressShown = $false
# The progress in course (Get-MmrProgressEta): its label, when and where it began.
$script:ProgressEta = $null

function Get-MmrIconSet {
    <# Icons of one console style. Symbols: only characters of the classic console fonts. #>
    param([Parameter(Mandatory = $true)][ValidateSet('Emoji', 'Symbols', 'Ascii')][string]$Style)

    $u = { param([int]$Code) [char]::ConvertFromUtf32($Code) }
    switch ($Style) {
        'Emoji' {
            return @{
                Logo = & $u 0x1F4EC; Ok = & $u 0x2705; Warn = (& $u 0x26A0) + [char]0xFE0F; Fail = & $u 0x274C; Info = & $u 0x1F539
                Skip = & $u 0x23E9; Key = & $u 0x1F511; Server = & $u 0x1F5A5; Shield = & $u 0x1F512; Archive = (& $u 0x1F5C4) + [char]0xFE0F
                Mail = & $u 0x1F4E8; People = & $u 0x1F465; User = & $u 0x1F464; File = & $u 0x1F4C4; Log = & $u 0x1F4DD
                Report = & $u 0x1F4CA; Done = & $u 0x1F389; Target = & $u 0x1F3AF; Search = & $u 0x1F50E; Clock = & $u 0x23F3
                Calendar = & $u 0x1F4C6; Folder = & $u 0x1F4C2; Trash = (& $u 0x1F5D1) + [char]0xFE0F; Cloud = (& $u 0x2601) + [char]0xFE0F; Filter = & $u 0x1F50D
            }
        }
        'Symbols' {
            return @{
                Logo = & $u 0x2666; Ok = & $u 0x221A; Warn = & $u 0x25B2; Fail = & $u 0x00D7; Info = & $u 0x2022
                Skip = & $u 0x00BB; Key = & $u 0x00A7; Server = & $u 0x2261; Shield = & $u 0x25CA; Archive = & $u 0x25A0
                Mail = '@'; People = & $u 0x2192; User = & $u 0x263A; File = & $u 0x25AC; Log = & $u 0x00B6
                Report = & $u 0x2261; Done = & $u 0x221A; Target = & $u 0x25D9; Search = & $u 0x25BA; Clock = & $u 0x25CB
                Calendar = & $u 0x25A1; Folder = & $u 0x2302; Trash = & $u 0x00D7; Cloud = & $u 0x2248; Filter = & $u 0x00A4
            }
        }
        default {
            return @{
                Logo = '*'; Ok = '+'; Warn = '!'; Fail = 'x'; Info = '-'; Skip = '>'; Key = 'k'; Server = '='; Shield = 'o'; Archive = '#'
                Mail = '@'; People = '&'; User = 'u'; File = '-'; Log = '='; Report = '='; Done = '*'; Target = 'o'; Search = '?'
                Clock = '~'; Calendar = '#'; Folder = '/'; Trash = 'x'; Cloud = '~'; Filter = 'f'
            }
        }
    }
}

function Get-MmrFrameSet {
    <# Rounded corners in modern terminals (emoji style), square corners elsewhere (present in every console font). #>
    param([Parameter(Mandatory = $true)][ValidateSet('Emoji', 'Symbols', 'Ascii')][string]$Style)

    if ($Style -eq 'Ascii') {
        return @{ TopLeft = [char]'+'; TopRight = [char]'+'; BottomLeft = [char]'+'; BottomRight = [char]'+'; Horizontal = [char]'-'; Vertical = [char]'|' }
    }
    if ($Style -eq 'Symbols') {
        return @{ TopLeft = [char]0x250C; TopRight = [char]0x2510; BottomLeft = [char]0x2514; BottomRight = [char]0x2518; Horizontal = [char]0x2500; Vertical = [char]0x2502 }
    }
    return @{ TopLeft = [char]0x256D; TopRight = [char]0x256E; BottomLeft = [char]0x2570; BottomRight = [char]0x256F; Horizontal = [char]0x2500; Vertical = [char]0x2502 }
}

$script:Icons = Get-MmrIconSet $script:IconStyle
$script:Frame = Get-MmrFrameSet $script:IconStyle
$script:IconPad = if ($script:IconStyle -eq 'Emoji') { ' ' } else { '  ' }
$script:IconWidth = if ($script:IconStyle -eq 'Emoji') { 2 } else { 1 }

function Get-MmrIcon { param([Parameter(Mandatory = $true)][string]$Name) return $script:Icons[$Name] + $script:IconPad }

function Format-MmrDuration {
    param([Parameter(Mandatory = $true)][double]$Seconds)

    $inv = [Globalization.CultureInfo]::InvariantCulture
    $t = [TimeSpan]::FromTicks([long]([Math]::Max(0.0, $Seconds) * 10000000))
    if ($t.TotalHours -ge 1) { return [string]::Format($inv, '{0} h {1:00} min', [int][Math]::Floor($t.TotalHours), $t.Minutes) }
    if ($t.TotalMinutes -ge 1) { return [string]::Format($inv, '{0} min {1:00} s', $t.Minutes, $t.Seconds) }
    return [string]::Format($inv, '{0:0.0} s', $t.TotalSeconds)
}

function Format-MmrText {
    <# Text cut to a width with an ellipsis, padded to the width. #>
    param([AllowEmptyString()][AllowNull()][string]$Text, [int]$Width)
    $t = [string]$Text -replace '[\r\n\t]+', ' '
    if ($Width -le 0) { return $t }
    if ($t.Length -gt $Width) { return $t.Substring(0, [Math]::Max(0, $Width - 1)) + [char]0x2026 }
    return $t.PadRight($Width)
}

function Send-MmrUi {
    <#
        Forwards a console line to the window while a window run is in progress: into the queue the window reads
        (background run, Ui.Queue) or to its sink (Ui.Sink).
    #>
    param([string]$Status, [string]$Text)
    $u = $script:Ui
    if (-not $u) { return }
    if ($u.Queue) { $u.Queue.Enqueue([string[]]@($Status, $Text)) }
    elseif ($u.Sink) { & $u.Sink $Status $Text }
}

function Start-MmrLog {
    <# Opens (or continues) today's log file and deletes the log files older than the retention. #>
    param([Parameter(Mandatory = $true)][string]$Directory, [int]$RetentionDays = 30)

    Stop-MmrLog
    [void][IO.Directory]::CreateDirectory($Directory)
    $script:LogPath = Join-Path $Directory ('MailboxMessageReport_{0:yyyyMMdd}.log' -f (Get-Date))
    $stream = [IO.FileStream]::new($script:LogPath, [IO.FileMode]::Append, [IO.FileAccess]::Write, [IO.FileShare]::ReadWrite)
    $writer = [IO.StreamWriter]::new($stream, [Text.UTF8Encoding]::new($false))
    $writer.AutoFlush = $true
    # Synchronized: the window and its background run write to the same log.
    $script:LogWriter = [IO.TextWriter]::Synchronized($writer)
    $limit = (Get-Date).AddDays(-$RetentionDays)
    Get-ChildItem -LiteralPath $Directory -Filter 'MailboxMessageReport_*.log' -File -ErrorAction SilentlyContinue |
        Where-Object LastWriteTime -lt $limit | Remove-Item -Force -ErrorAction SilentlyContinue
    return $script:LogPath
}

function Stop-MmrLog {
    if ($script:LogWriter) { $script:LogWriter.Dispose(); $script:LogWriter = $null }
}

function Write-MmrLog {
    <# One line in the log file only. The log never contains colours, icons, tokens or secrets. #>
    param(
        [ValidateSet('INFO', 'OK', 'WARN', 'ERROR', 'STEP')][string]$Level = 'INFO',
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Message
    )
    if ($script:LogWriter) { $script:LogWriter.WriteLine(('{0:yyyy-MM-ddTHH:mm:ss.fffzzz} [{1,-5}] {2}' -f (Get-Date), $Level, $Message)) }
}

function Get-MmrConsoleWidth {
    try { $w = [Console]::WindowWidth; if ($w -ge 40) { return $w } } catch { }
    return 120
}

function Clear-MmrProgress {
    <# Ends the live progress line (if any) so that the next line starts on a new row. #>
    if ($script:ProgressShown) {
        [Console]::Write("`r" + (' ' * [Math]::Max(10, (Get-MmrConsoleWidth) - 1)) + "`r")
        $script:ProgressShown = $false
    }
}

function Write-MmrBanner {
    <# Title card at the start of an execution, followed by the context rows (label -> @(Icon, Text)). #>
    param(
        [Parameter(Mandatory = $true)][string]$Title,
        [string]$Subtitle,
        [System.Collections.Specialized.OrderedDictionary]$Details
    )

    Write-MmrLog 'STEP' "=== $Title v$($script:ToolVersion) ==="
    if ($Details) {
        foreach ($key in $Details.Keys) {
            $v = $Details[$key]
            Write-MmrLog 'INFO' ('{0}: {1}' -f $key, $(if ($v -is [array]) { $v[1] } else { $v }))
        }
    }
    if ($script:Quiet) { return }
    $C = $script:C; $F = $script:Frame; $width = 78
    $right = "v$($script:ToolVersion) $($script:Dot) Nicolas Fabert"
    $left = "  $($script:Icons.Logo)  $Title"
    $gap = [Math]::Max(1, $width - ($left.Length - $script:Icons.Logo.Length + $script:IconWidth) - $right.Length - 2)
    Write-Host ''
    Write-Host ('  {0}{1}{2}{3}{4}' -f $C.Accent, $F.TopLeft, [string]::new($F.Horizontal, $width), $F.TopRight, $C.Reset)
    Write-Host ('  {0}{1}{2}{3}{4}{5}{6}{7}{8}{9}{10}{11}' -f $C.Accent, $F.Vertical, $C.Reset, $C.Bold, $left, $C.Reset, [string]::new(' ', $gap), $C.Dim, $right, '  ', ($C.Accent + $F.Vertical), $C.Reset)
    if ($Subtitle) {
        Write-Host ('  {0}{1}{2}{3}{4}{5}{0}{6}{2}' -f $C.Accent, $F.Vertical, $C.Reset, $C.Dim, (Format-MmrText "     $Subtitle" $width), $C.Reset, $F.Vertical)
    }
    Write-Host ('  {0}{1}{2}{3}{4}' -f $C.Accent, $F.BottomLeft, [string]::new($F.Horizontal, $width), $F.BottomRight, $C.Reset)
    if ($Details) {
        foreach ($key in $Details.Keys) {
            $value = $Details[$key]
            $icon, $text = if ($value -is [array]) { (Get-MmrIcon $value[0]), $value[1] } else { '   ', $value }
            Write-Host ('     {0}{1}{2,-11}{3} {4}' -f $icon, $C.Dim, $key, $C.Reset, $text)
        }
    }
}

function Write-MmrStep {
    <# Step header with a coloured number pill and an icon:  ─ 3/6 ─ 🔎  Search #>
    param(
        [Parameter(Mandatory = $true)][int]$Number,
        [Parameter(Mandatory = $true)][int]$Total,
        [Parameter(Mandatory = $true)][string]$Title,
        [string]$Icon = 'Info'
    )

    Write-MmrLog 'STEP' "[$Number/$Total] $Title"
    $script:ProgressEta = $null
    Send-MmrUi 'Step' "[$Number/$Total] $Title"
    if ($script:Quiet) { return }
    Clear-MmrProgress
    $C = $script:C
    Write-Host ''
    Write-Host ('  {0} {1}/{2} {3} {4}{5}{6}{3}' -f $C.AccentBg, $Number, $Total, $C.Reset, (Get-MmrIcon $Icon), $C.Bold, $Title)
}

function Write-MmrItem {
    <# One indented result line with a status icon, also written to the log and to the window. #>
    param(
        [ValidateSet('Ok', 'Warn', 'Fail', 'Info', 'Skip')][string]$Status = 'Info',
        [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Text,
        [string]$Icon
    )

    $level = @{ Ok = 'OK'; Warn = 'WARN'; Fail = 'ERROR'; Info = 'INFO'; Skip = 'INFO' }[$Status]
    Write-MmrLog $level $Text
    Send-MmrUi $Status $Text
    if ($script:Quiet) { return }
    Clear-MmrProgress
    $color = @{ Ok = $script:C.Green; Warn = $script:C.Yellow; Fail = $script:C.Red; Info = ''; Skip = $script:C.Dim }[$Status]
    $symbol = Get-MmrIcon $(if ($Icon) { $Icon } else { $Status })
    $textColor = if ($Status -in 'Warn', 'Fail', 'Skip') { $color } else { '' }
    Write-Host ('      {0}{1}{2}{3}{4}{2}' -f $color, $symbol, $script:C.Reset, $textColor, $Text)
}

function Format-MmrTimeLeft {
    <#
        The time left of a progress, rounded as a person would say it: a few seconds, about 25 s (5 s steps under a
        minute), about 1 min 30 s (10 s steps under 5 minutes), about 12 min, about 1 h 05 min.
    #>
    param([Parameter(Mandatory = $true)][double]$Seconds)

    $inv = [Globalization.CultureInfo]::InvariantCulture
    if ($Seconds -lt 10) { return 'a few seconds left' }
    $away = [MidpointRounding]::AwayFromZero
    $r = [int]$(if ($Seconds -lt 60) { [Math]::Ceiling($Seconds / 5) * 5 } elseif ($Seconds -lt 300) { [Math]::Round($Seconds / 10, $away) * 10 } else { [Math]::Round($Seconds / 60, $away) * 60 })
    if ($r -lt 60) { return [string]::Format($inv, 'about {0} s left', $r) }
    $h = [int][Math]::Floor($r / 3600); $m = [int][Math]::Floor(($r % 3600) / 60); $s = $r % 60
    if ($h) { return [string]::Format($inv, 'about {0} h {1:00} min left', $h, $m) }
    if ($s) { return [string]::Format($inv, 'about {0} min {1:00} s left', $m, $s) }
    return [string]::Format($inv, 'about {0} min left', $m)
}

function Get-MmrProgressEta {
    <#
        Time left of the progress in course, from its speed since it began; empty until it can be told (2 s and
        2 % of progress since its first value). Another label (the counts aside), or a value going back, is a
        new progress. A step starts with none (Write-MmrStep). -Now: tests.
    #>
    param([Parameter(Mandatory = $true)][double]$Fraction, [AllowEmptyString()][string]$Text, [datetime]$Now = [datetime]::UtcNow)

    $key = $Text -replace '[\d\s,.\u00A0\u202F/]+', ''
    $s = $script:ProgressEta
    if (-not $s -or $s.Key -ne $key -or $Fraction -lt $s.Last) {
        $script:ProgressEta = @{ Key = $key; Start = $Now; From = $Fraction; Last = $Fraction; Left = -1.0; At = $Now }
        return ''
    }
    $s.Last = $Fraction
    $done = $Fraction - $s.From
    $elapsed = ($Now - $s.Start).TotalSeconds
    if ($Fraction -ge 1 -or $done -lt 0.02 -or $elapsed -lt 2) { return '' }
    $left = $elapsed / $done * (1 - $Fraction)
    # Graph answers come in bursts (16 calls in flight): half the new figure, half the last one brought forward.
    if ($s.Left -ge 0) { $left = 0.5 * $left + 0.5 * [Math]::Max(0.0, $s.Left - ($Now - $s.At).TotalSeconds) }
    $s.Left = $left; $s.At = $Now
    return Format-MmrTimeLeft $left
}

function Write-MmrProgress {
    <#
        Live progress line, rewritten in place (interactive console); sent to the window during a window run
        (fraction|text|time left).
              ⏳  ███████░░░░░  58%  1,077/1,858 folders read · 254,310 messages · about 40 s left
    #>
    param([Parameter(Mandatory = $true)][double]$Fraction, [Parameter(Mandatory = $true)][string]$Text)

    $left = Get-MmrProgressEta -Fraction $Fraction -Text $Text
    Send-MmrUi 'Progress' ('{0}|{1}|{2}' -f $Fraction.ToString('0.000', [Globalization.CultureInfo]::InvariantCulture), $Text, $left)
    if ($script:Quiet -or [Console]::IsOutputRedirected) { return }
    $C = $script:C
    $percent = [int][Math]::Floor(100 * [Math]::Min(1.0, [Math]::Max(0.0, $Fraction)))
    $filled = [int][Math]::Round(12 * $percent / 100.0)
    $bar = $C.Accent + [string]::new([char]0x2588, $filled) + $C.Dim + [string]::new([char]0x2591, 12 - $filled) + $C.Reset
    $line = if ($left) { "$Text $($script:Dot) $left" } else { $Text }
    $plain = Format-MmrText $line ([Math]::Max(10, (Get-MmrConsoleWidth) - 30))
    [Console]::Write(("`r      {0}{1} {2,3}%  {3}{4}{5}" -f (Get-MmrIcon 'Clock'), $bar, $percent, $C.Dim, $plain.TrimEnd(), $C.Reset))
    $script:ProgressShown = $true
}

function Write-MmrTable {
    <#
        Aligned table, one row per object, with a status icon in front of each row.
        Columns: @{ Name = 'Header'; Property = 'PropertyName'; Width = 20; Align = 'Right' } - Width 0 = the rest of the console.
        StatusProperty: Ok | Warn | Fail | Info | Skip (colour and icon of the row).
    #>
    param(
        [Parameter(Mandatory = $true)][object[]]$Columns,
        [AllowEmptyCollection()][AllowNull()][object[]]$Rows,
        [string]$StatusProperty = 'Status',
        [int]$Indent = 6,
        [int]$MaxWidth = 170
    )

    if (-not $Rows -or -not $Rows.Count) { return }
    foreach ($row in $Rows) { Write-MmrLog 'INFO' (($Columns | ForEach-Object { "$($_.Name)=$([string]$row.($_.Property))" }) -join ' | ') }
    if ($script:Quiet) { return }
    Clear-MmrProgress
    $C = $script:C
    $consoleWidth = [Math]::Min($MaxWidth, (Get-MmrConsoleWidth) - 1)
    if ($consoleWidth -lt 80) { $consoleWidth = 120 }
    $fixed = [int](($Columns | ForEach-Object { [int]$_['Width'] } | Measure-Object -Sum).Sum) + 2 * $Columns.Count
    $last = [Math]::Max(20, $consoleWidth - $Indent - 3 - $fixed)
    $pad = ' ' * $Indent
    $cell = {
        param($col, $text)
        $w = if ([int]$col['Width']) { [int]$col['Width'] } else { $last }
        if ($col['Align'] -eq 'Right') { (Format-MmrText $text $w).Trim().PadLeft($w) } else { Format-MmrText $text $w }
    }
    $header = ($Columns | ForEach-Object { & $cell $_ $_['Name'] }) -join '  '
    Write-Host ('{0}{1}{2}{3}{4}' -f $pad, $C.Dim, (' ' * ($script:IconWidth + $script:IconPad.Length)), $header.TrimEnd(), $C.Reset)
    foreach ($row in $Rows) {
        $status = [string]$row.$StatusProperty
        if ($status -notin 'Ok', 'Warn', 'Fail', 'Info', 'Skip') { $status = 'Info' }
        $color = @{ Ok = $C.Green; Warn = $C.Yellow; Fail = $C.Red; Info = $C.Blue; Skip = $C.Dim }[$status]
        $cells = foreach ($col in $Columns) { & $cell $col ([string]$row.($col.Property)) }
        $textColor = if ($status -eq 'Skip') { $C.Dim } else { '' }
        Write-Host ('{0}{1}{2}{3}{4}{5}{3}' -f $pad, $color, (Get-MmrIcon $status), $C.Reset, $textColor, (($cells -join '  ').TrimEnd()))
    }
}

function Write-MmrSummary {
    <# Final summary card (label -> @(Icon, Text)). #>
    param(
        [Parameter(Mandatory = $true)][string]$Title,
        [Parameter(Mandatory = $true)][System.Collections.Specialized.OrderedDictionary]$Values,
        [ValidateSet('Ok', 'Warn', 'Fail')][string]$Status = 'Ok'
    )

    foreach ($key in $Values.Keys) {
        $v = $Values[$key]
        Write-MmrLog 'INFO' ('Summary - {0}: {1}' -f $key, $(if ($v -is [array]) { $v[1] } else { $v }))
    }
    if ($script:Quiet) { return }
    Clear-MmrProgress
    $C = $script:C; $F = $script:Frame; $width = 78
    $color = @{ Ok = $C.Green; Warn = $C.Yellow; Fail = $C.Red }[$Status]
    $icon = $script:Icons[@{ Ok = 'Done'; Warn = 'Warn'; Fail = 'Fail' }[$Status]]
    $head = " $icon  $Title "
    $rest = [Math]::Max(2, $width - 1 - ($head.Length - $icon.Length + $script:IconWidth))
    Write-Host ''
    Write-Host ('  {0}{1}{2}{3}{4}{0}{5}{6}{7}' -f $color, $F.TopLeft, $F.Horizontal, $C.Bold, $head, ($C.Reset + $color), ([string]::new($F.Horizontal, $rest) + $F.TopRight), $C.Reset)
    foreach ($key in $Values.Keys) {
        $value = $Values[$key]
        $rowIcon, $text = if ($value -is [array]) { (Get-MmrIcon $value[0]), $value[1] } else { '   ', $value }
        Write-Host ('    {0}{1}{2,-10}{3} {4}' -f $rowIcon, $C.Dim, $key, $C.Reset, $text)
    }
    Write-Host ('  {0}{1}{2}{3}{4}' -f $color, $F.BottomLeft, [string]::new($F.Horizontal, $width), $F.BottomRight, $C.Reset)
    Write-Host ''
}


function Format-MmrMailboxList {
    <# The mailboxes in one line: all of them up to 3, else the first ones and the count (and the file). #>
    param([object[]]$Mailboxes, [string]$File)
    $list = @($Mailboxes | ForEach-Object { if ($_.ArchiveGuid) { "$($_.Address) (ArchiveGuid)" } else { $_.Address } })
    $text = if ($list.Count -le 3) { $list -join ', ' } else { '{0} mailboxes ({1}, ...)' -f $list.Count, (($list | Select-Object -First 2) -join ', ') }
    if ($File) { $text += " $($script:Dot) file $([IO.Path]::GetFileName($File))" }
    return $text
}

function Get-MmrWhereText {
    <# Where the messages are read, in words. #>
    param([string[]]$Location, [bool]$RecoverableItems)
    $text = if ($Location -contains 'Primary' -and $Location -contains 'Archive') { 'primary mailbox and archive' } elseif ($Location -contains 'Archive') { 'archive only' } else { 'primary mailbox only' }
    if ($RecoverableItems) { $text += ', Recoverable Items included' }
    return $text
}

function Write-MmrRunBanner {
    <# Title card of a command-line run: mailboxes, messages, where, report, application and log. #>
    param(
        [Parameter(Mandatory = $true)][hashtable]$Settings,
        [Parameter(Mandatory = $true)][pscustomobject]$Request,
        [string]$LogPath
    )

    $dot = $script:Dot
    $banner = [ordered]@{}
    $banner['Mailboxes'] = @('User', (Format-MmrMailboxList -Mailboxes $Request.Mailboxes -File $Request.MailboxFile))
    $what = if ($null -ne $Request.Start -or $null -ne $Request.End) {
        "received $(if ($null -ne $Request.Start) { "from $(Format-MmrDate $Request.Start $Settings.TimeZone)" }) $(if ($null -ne $Request.End) { "to $(Format-MmrDate $Request.End $Settings.TimeZone -PeriodEnd)" })".Trim() -replace '\s+', ' '
    } else { 'every message, whatever its date' }
    if (@($Request.Subject).Count) { $what += " $dot subject contains $((@($Request.Subject) | ForEach-Object { "'$_'" }) -join ' or ')" }
    if (-not $Request.Recipients) { $what += " $dot without recipients (To, Cc, Bcc)" }
    $banner['Messages'] = @('Mail', $what)
    $where = Get-MmrWhereText -Location $Request.Location -RecoverableItems $Request.RecoverableItems
    if (@($Request.ExcludeFolders).Count) { $where += " $dot left out: $(@($Request.ExcludeFolders) -join ', ')" }
    $banner['Where'] = @('Archive', $where)
    $banner['Tenant'] = @('Cloud', $(if ($Settings.Organization) { "$($Settings.Organization) $dot $($Settings.TenantId)" } else { $Settings.TenantId }))
    $banner['App'] = @('Key', "$($Settings.AppId) $dot $(if ($Settings.AuthMode -eq 'Certificate') { "certificate $($Settings.CertificateThumbprint)" } else { "client secret (`$env:$($Settings.ClientSecretVariable) or prompt)" })")
    $banner['Report'] = @('Report', "$($Settings.OutputPath) $dot $(@($Request.Formats) -join ', ') $dot $(switch ($Request.Layout) { 'PerMailbox' { 'one report per mailbox' } 'Both' { 'global and per mailbox' } default { 'global' } })")
    if ($LogPath) { $banner['Log'] = @('Log', $LogPath) }
    Write-MmrBanner -Title 'Mailbox Message Report' -Subtitle "Exchange Online $dot messages of the primary mailbox and the archive, through Microsoft Graph" -Details $banner
}

function Write-MmrMailboxTable {
    <# The mailboxes of a result, one line each: status, archive, folders, messages (primary, archive, Recoverable Items). #>
    param([Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Mailboxes, [int]$Max = 200)

    $rows = foreach ($m in ($Mailboxes | Select-Object -First $Max)) {
        [pscustomobject]@{
            Status      = switch ($m.Status) { 'Read' { 'Ok' } 'Partial' { 'Warn' } 'Not read' { 'Fail' } default { 'Info' } }
            Mailbox     = $m.Address
            Archive     = switch ($m.Archive) { 'Yes' { if ($m.ArchiveSource -eq 'File') { 'Yes (list)' } else { 'Yes' } } 'No' { 'No' } default { '?' } }
            Folders     = '{0:N0}{1}' -f $m.FoldersRead, $(if ($m.FoldersFailed) { " ($($m.FoldersFailed) err.)" } else { '' })
            Primary     = '{0:N0}' -f $m.PrimaryMessages
            InArchive   = '{0:N0}' -f $m.ArchiveMessages
            Recoverable = '{0:N0}' -f $m.RecoverableMessages
            Notes       = if ($m.State -ne 'Ok') { $m.Detail } else { $m.Notes }
        }
    }
    $columns = @(
        @{ Name = 'Mailbox'; Property = 'Mailbox'; Width = 34 }
        @{ Name = 'Archive'; Property = 'Archive'; Width = 10 }
        @{ Name = 'Folders'; Property = 'Folders'; Width = 12; Align = 'Right' }
        @{ Name = 'Primary'; Property = 'Primary'; Width = 10; Align = 'Right' }
        @{ Name = 'Archive'; Property = 'InArchive'; Width = 10; Align = 'Right' }
        @{ Name = 'Recov.'; Property = 'Recoverable'; Width = 8; Align = 'Right' }
        @{ Name = 'Notes'; Property = 'Notes'; Width = 0 }
    )
    Write-MmrTable -Rows @($rows) -Columns $columns
    if ($Mailboxes.Count -gt $Max) { Write-MmrItem Info ('... and {0:N0} more mailboxes: see the report (Mailboxes tab, Mailboxes.csv).' -f ($Mailboxes.Count - $Max)) }
}

function Write-MmrRunSummary {
    <# Final card of a command-line run: status, mailboxes, folders, messages, report, log and what to do next. #>
    param(
        [Parameter(Mandatory = $true)][pscustomobject]$Result,
        [string]$ReportText,
        [string]$LogPath
    )

    $dot = $script:Dot
    $n = $Result.Counts
    $values = [ordered]@{}
    $values['Status'] = @($(switch ($Result.Status) { 'Completed' { 'Ok' } 'Failed' { 'Fail' } default { 'Warn' } }), "$($Result.Status) $dot read only: nothing is changed in the mailboxes")
    $values['Mailboxes'] = @('People', ('{0:N0} {1} {2:N0} read {1} {3:N0} in part {1} {4:N0} not read {1} {5:N0} with an archive' -f $n.Mailboxes, $dot, $n.MailboxesRead, $n.MailboxesPartial, $n.MailboxesNotRead, $n.WithArchive))
    $values['Folders'] = @('Folder', ('{0:N0} {1} {2:N0} read {1} {3:N0} empty {1} {4:N0} left out {1} {5:N0} in error' -f $n.Folders, $dot, $n.FoldersRead, $n.FoldersEmpty, $n.FoldersExcluded, $n.FoldersFailed))
    $values['Messages'] = @('Mail', ('{0:N0} {1} {2:N0} primary {1} {3:N0} archive{4}' -f $n.Messages, $dot, $n.PrimaryMessages, $n.ArchiveMessages, $(if ($Result.Request.RecoverableItems) { " $dot $('{0:N0}' -f $n.RecoverableMessages) Recoverable Items" })))
    if ($Result.Error) { $values['First issue'] = @('Fail', $Result.Error) }
    $values['Duration'] = @('Clock', (Format-MmrDuration $Result.DurationSeconds))
    if ($ReportText) { $values['Report'] = @('Report', $ReportText) }
    if ($LogPath) { $values['Log'] = @('Log', $LogPath) }
    $values['Next'] = @('Info', $(switch ($Result.Status) {
                'Completed' { if ($n.Messages) { 'Open the HTML report, or the CSV file in Excel (every message, separator of the configuration).' } else { 'No message: widen the period, check the subject, or add the archive (-Location Primary, Archive).' } }
                'Failed' { 'Read the first issue above and the log.' }
                default { 'Open the report: the Mailboxes and Folders tabs say what was not read and why. Running the same command again reads everything again.' }
            }))
    $title = switch ($Result.Status) { 'Completed' { 'Report finished' } 'Failed' { 'Run failed' } default { 'Finished with warnings' } }
    $card = switch ($Result.Status) { 'Completed' { 'Ok' } 'Failed' { 'Fail' } default { 'Warn' } }
    Write-MmrSummary -Title $title -Values $values -Status $card
}
