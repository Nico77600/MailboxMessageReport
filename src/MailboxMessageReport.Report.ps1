<#
.SYNOPSIS
    Mailbox Message Report - report files (dot-sourced by MailboxMessageReport.psm1).

.DESCRIPTION
    One folder per run, <FilePrefix>_<yyyyMMdd-HHmmss>, with:
      <prefix>-Messages.csv     Global layout: one row per message of every mailbox, in the order of the report
                                (mailbox, primary mailbox then archive, Recoverable Items last, folder path, newest first)
      <prefix>-Mailboxes.csv    one row per mailbox: archive, folders, messages in the primary mailbox / archive /
                                Recoverable Items, why a mailbox was not read
      <prefix>-Folders.csv      one row per folder: location, path, items, messages found, status
      <prefix>-Summary.json     the whole result (request, mailboxes, folders, counts), for scripts
      <prefix>.html             self-contained dashboard (templates\Report.template.html); PerMailbox layout: the
                                summary, with a link to the report of each mailbox
      Mailboxes\<prefix>-<address>.csv / .html   PerMailbox layout: the messages of one mailbox
    The messages come from the part files written while reading (one per folder, one JSON array per message and per
    line): they are merged in order into the CSV files and the HTML reports (compiled: MailboxMessageReportNative.Merge),
    then deleted. A HTML report holds the first Report.HtmlMaxMessages messages (the CSV file holds them all).
    CSV files: UTF-8 with BOM, configurable delimiter, text cells starting with = + - @ are prefixed with an apostrophe
    (no formula injection when opened in Excel). No token or secret is ever part of the result.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.0.0
#>

$script:ReportColumns = [ordered]@{
    Mailboxes = @('Input', 'Address', 'DisplayName', 'Status', 'State', 'Detail', 'Archive', 'ArchiveSource', 'ArchiveGuid', 'PrimaryFolders', 'ArchiveFolders', 'FoldersRead', 'FoldersFailed', 'PrimaryMessages', 'ArchiveMessages', 'RecoverableMessages', 'Messages', 'Notes')
    Folders   = @('Mailbox', 'MailboxName', 'Location', 'RecoverableItems', 'Path', 'Name', 'TotalItems', 'Messages', 'Status', 'Detail')
}

function New-MmrRunFolder {
    <# New folder of a run under OutputPath: <Prefix>_<yyyyMMdd-HHmmss>. #>
    param([Parameter(Mandatory = $true)][string]$OutputPath, [string]$Prefix = 'MailboxMessageReport')
    $base = Join-Path $OutputPath ('{0}_{1}' -f $Prefix, (Get-Date).ToString('yyyyMMdd-HHmmss'))
    $path = $base
    $n = 2
    while (Test-Path -LiteralPath $path) { $path = "$base-$n"; $n++ }
    [void][IO.Directory]::CreateDirectory($path)
    return $path
}

function Get-MmrSafeFileName {
    <# A text usable in a file name (an address keeps its @ and dots). #>
    param([Parameter(Mandatory = $true)][string]$Text)
    $invalid = [IO.Path]::GetInvalidFileNameChars()
    $sb = [Text.StringBuilder]::new()
    foreach ($ch in $Text.ToCharArray()) { [void]$sb.Append($(if ($invalid -contains $ch -or [char]::IsWhiteSpace($ch)) { '_' } else { $ch })) }
    return $sb.ToString()
}

function Write-MmrTableCsv {
    <# A CSV file of PowerShell objects: UTF-8 with BOM, the columns given, cells neutralised (compiled). #>
    param([AllowEmptyCollection()][object[]]$Rows, [Parameter(Mandatory = $true)][string[]]$Columns, [Parameter(Mandatory = $true)][string]$Path, [string]$Delimiter = ';')
    $target = [MailboxMessageReportNative.CsvTarget]::new($Path, $Columns, $Delimiter)
    try {
        foreach ($r in @($Rows)) {
            $cells = [string[]]::new($Columns.Count)
            for ($i = 0; $i -lt $Columns.Count; $i++) {
                $v = Get-MmrProperty $r $Columns[$i]
                $cells[$i] = if ($v -is [bool]) { if ($v) { 'Yes' } else { 'No' } } elseif ($v -is [array]) { $v -join ' | ' } else { [string]$v }
            }
            $target.Write($cells)
        }
    }
    finally { $target.Dispose() }
}

function ConvertTo-MmrEmbeddedJson {
    <# JSON safe inside a <script type="application/json"> block. #>
    param([AllowNull()][object]$Value)
    $json = ConvertTo-Json -InputObject $Value -Depth 10 -Compress
    if ([string]::IsNullOrEmpty($json)) { $json = 'null' }
    return $json.Replace('<', '\u003c').Replace('>', '\u003e').Replace('&', '\u0026')
}

function Get-MmrStatusText {
    <# The status of a run or of a mailbox, in words. #>
    param([string]$Status)
    switch ($Status) { 'Completed' { 'Completed' } 'Warning' { 'Finished with warnings' } 'Failed' { 'Failed' } default { $Status } }
}

function Write-MmrHtml {
    <#
        One HTML report from the template: the summary, the mailboxes, the folders, and the messages kept in a row buffer
        (written by the compiled helper: the JSON of thousands of messages never goes through a PowerShell string).
    #>
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Title,
        [Parameter(Mandatory = $true)][System.Collections.IDictionary]$Summary,
        [AllowEmptyCollection()][object[]]$Mailboxes,
        [AllowEmptyCollection()][object[]]$Folders,
        [MailboxMessageReportNative.RowBuffer]$Messages
    )
    $mbx = @($Mailboxes | Select-Object -Property ($script:ReportColumns.Mailboxes + 'Html', 'Csv'))
    $fld = @($Folders | Select-Object -Property $script:ReportColumns.Folders)
    $markers = [string[]]@('{{TITLE}}', '{{SUMMARY_JSON}}', '{{MAILBOXES_JSON}}', '{{FOLDERS_JSON}}')
    $values = [string[]]@([Net.WebUtility]::HtmlEncode($Title), (ConvertTo-MmrEmbeddedJson $Summary), (ConvertTo-MmrEmbeddedJson $mbx), (ConvertTo-MmrEmbeddedJson $fld))
    [MailboxMessageReportNative.Merge]::WriteHtml((Join-Path $script:ToolRoot 'templates\Report.template.html'), $Path, $markers, $values, $Messages, '{{MESSAGES_JSON}}')
}

function Get-MmrHtmlSummary {
    <# The summary given to a HTML report: the run, or one mailbox of it (-Mailbox). #>
    param([Parameter(Mandatory = $true)][pscustomobject]$Result, $Mailbox, [System.Collections.IDictionary]$Counts, [long]$Shown, [long]$Total, [string]$Layout, [bool]$MessagesHere)
    $summary = [ordered]@{}
    foreach ($key in 'Tool', 'Version', 'Status', 'Error', 'StartedUtc', 'CompletedUtc', 'DurationSeconds', 'Request', 'Tenant', 'Organization', 'AppId', 'AppName', 'Roles', 'Warnings') {
        $summary[$key] = $null
        $prop = $Result.PSObject.Properties[$key]
        if ($prop) { $summary[$key] = $prop.Value }
    }
    $summary.Counts = $Counts
    $summary.Scope = if ($Mailbox) { 'Mailbox' } else { 'Run' }
    $summary.Mailbox = if ($Mailbox) { [ordered]@{ Address = $Mailbox.Address; DisplayName = $Mailbox.DisplayName; Status = $Mailbox.Status; Archive = $Mailbox.Archive; Detail = $Mailbox.Detail } } else { $null }
    $summary.Layout = $Layout
    $summary.MessagesHere = $MessagesHere
    $summary.MessagesShown = $Shown
    $summary.MessagesTotal = $Total
    $summary.Columns = [MailboxMessageReportNative.Columns]::Messages
    $summary.StatusText = Get-MmrStatusText ([string]$Result.Status)
    $summary.GeneratedText = Format-MmrDate ([datetime]::UtcNow) ([string]$Result.Request.TimeZone)
    return $summary
}

function Export-MmrReport {
    <#
    .SYNOPSIS
        Writes the report of a run in its folder: the messages merged from the part files (CSV, HTML), the mailboxes,
        the folders and Summary.json. The part files are deleted.
    .PARAMETER Layout
        Global (one file for every mailbox), PerMailbox (one file per mailbox, under Mailboxes\) or Both.
    .PARAMETER PreviewMessages
        How many messages to keep for the window (the first ones of the report).
    .OUTPUTS
        @{ Directory; Files (ordered: Html, Messages, Mailboxes, Folders, Summary); MailboxFiles; Preview (RowBuffer) }.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][pscustomobject]$Result,
        [Parameter(Mandatory = $true)][string]$Directory,
        [string]$Prefix = 'MailboxMessageReport',
        [ValidateSet('Csv', 'Html')][string[]]$Formats = @('Csv', 'Html'),
        [ValidateSet('Global', 'PerMailbox', 'Both')][string]$Layout = 'Global',
        [ValidateSet(';', ',', "`t")][string]$Delimiter = ';',
        [int]$HtmlMaxMessages = 20000,
        [int]$PreviewMessages = 0,
        [string]$PartsPath
    )

    [void][IO.Directory]::CreateDirectory($Directory)
    $global = $Layout -in 'Global', 'Both'
    $perMailbox = $Layout -in 'PerMailbox', 'Both'
    $csv = $Formats -contains 'Csv'
    $html = $Formats -contains 'Html'
    $files = [ordered]@{}
    $mailboxFiles = 0
    $byMailbox = @{}
    foreach ($f in @($Result.Folders)) {
        if (-not $byMailbox.ContainsKey($f.Mailbox)) { $byMailbox[$f.Mailbox] = [Collections.Generic.List[object]]::new() }
        $byMailbox[$f.Mailbox].Add($f)
    }

    # ---- the messages: part files merged in the order of the report ---------------------------------------------
    $globalCsv = $null
    $globalBuffer = if ($global -and $html) { [MailboxMessageReportNative.RowBuffer]::new($HtmlMaxMessages) } else { $null }
    $preview = [MailboxMessageReportNative.RowBuffer]::new([Math]::Max(0, $PreviewMessages))
    try {
        if ($global -and $csv) { $files.Messages = Join-Path $Directory "$Prefix-Messages.csv"; $globalCsv = [MailboxMessageReportNative.CsvTarget]::new($files.Messages, [MailboxMessageReportNative.Columns]::Messages, $Delimiter) }
        $sub = Join-Path $Directory 'Mailboxes'
        if ($perMailbox) { [void][IO.Directory]::CreateDirectory($sub) }
        $done = 0
        foreach ($m in @($Result.Mailboxes)) {
            $done++
            Write-MmrProgress ($done / [Math]::Max(1, @($Result.Mailboxes).Count)) ('{0:N0}/{1:N0} mailboxes written' -f $done, @($Result.Mailboxes).Count)
            # A mailbox not read has no report of its own: the summary says why.
            $own = $perMailbox -and $m.State -eq 'Ok'
            $mailboxCsv = $null
            $mailboxBuffer = if ($own -and $html) { [MailboxMessageReportNative.RowBuffer]::new($HtmlMaxMessages) } else { $null }
            $name = Get-MmrSafeFileName $m.Address
            try {
                if ($own -and $csv) { $m.Csv = "Mailboxes/$Prefix-$name.csv"; $mailboxCsv = [MailboxMessageReportNative.CsvTarget]::new((Join-Path $sub "$Prefix-$name.csv"), [MailboxMessageReportNative.Columns]::Messages, $Delimiter) }
                $targets = [MailboxMessageReportNative.CsvTarget[]]@(@($globalCsv, $mailboxCsv) | Where-Object { $_ })
                $buffers = [MailboxMessageReportNative.RowBuffer[]]@(@($globalBuffer, $mailboxBuffer, $preview) | Where-Object { $_ })
                $folders = if ($byMailbox.ContainsKey($m.Address)) { $byMailbox[$m.Address] } else { @() }
                foreach ($f in $folders) {
                    if ($f.Messages -gt 0) { foreach ($part in @($f.Parts)) { [void][MailboxMessageReportNative.Merge]::AppendPart($part, $targets, $buffers) } }
                }
            }
            finally { if ($mailboxCsv) { $mailboxCsv.Dispose(); $mailboxFiles++ } }
            if ($mailboxBuffer) {
                $m.Html = "Mailboxes/$Prefix-$name.html"
                $counts = Get-MmrCounts -Mailboxes @($m) -Folders @($folders)
                $summary = Get-MmrHtmlSummary -Result $Result -Mailbox $m -Counts $counts -Shown $mailboxBuffer.Lines.Count -Total $mailboxBuffer.Seen -Layout $Layout -MessagesHere $true
                $title = "Mailbox Message Report | $(if ($m.DisplayName) { $m.DisplayName } else { $m.Address })"
                Write-MmrHtml -Path (Join-Path $sub "$Prefix-$name.html") -Title $title -Summary $summary -Mailboxes @($m) -Folders @($folders) -Messages $mailboxBuffer
                if (-not $csv) { $mailboxFiles++ }
            }
        }
    }
    finally { if ($globalCsv) { $globalCsv.Dispose() } }

    # ---- the mailboxes, the folders, the summary --------------------------------------------------------------------
    if ($csv) {
        $files.Mailboxes = Join-Path $Directory "$Prefix-Mailboxes.csv"
        Write-MmrTableCsv -Rows @($Result.Mailboxes) -Columns $script:ReportColumns.Mailboxes -Path $files.Mailboxes -Delimiter $Delimiter
        $files.Folders = Join-Path $Directory "$Prefix-Folders.csv"
        Write-MmrTableCsv -Rows @($Result.Folders) -Columns $script:ReportColumns.Folders -Path $files.Folders -Delimiter $Delimiter
    }
    if ($html) {
        $files.Html = Join-Path $Directory "$Prefix.html"
        $shown = if ($globalBuffer) { $globalBuffer.Lines.Count } else { 0 }
        $total = if ($globalBuffer) { $globalBuffer.Seen } else { [long]$Result.Counts.Messages }
        $summary = Get-MmrHtmlSummary -Result $Result -Counts $Result.Counts -Shown $shown -Total $total -Layout $Layout -MessagesHere ($null -ne $globalBuffer)
        $names = @($Result.Mailboxes | ForEach-Object { if ($_.DisplayName) { $_.DisplayName } else { $_.Address } })
        $title = "Mailbox Message Report | $(if ($names.Count -le 3) { $names -join ', ' } else { "$($names.Count) mailboxes" })"
        Write-MmrHtml -Path $files.Html -Title $title -Summary $summary -Mailboxes @($Result.Mailboxes) -Folders @($Result.Folders) -Messages $globalBuffer
    }
    $files.Summary = Join-Path $Directory "$Prefix-Summary.json"
    $Result.RunFolder = $Directory
    $Result.Files = $files
    $data = [ordered]@{}
    foreach ($p in $Result.PSObject.Properties) { $data[$p.Name] = $p.Value }
    # The part files are gone once the report is written.
    $data.Folders = @($Result.Folders | Select-Object -Property * -ExcludeProperty Parts, SlicesDone)
    [IO.File]::WriteAllText($files.Summary, (ConvertTo-Json -InputObject $data -Depth 8), [Text.UTF8Encoding]::new($false))

    if ($PartsPath -and (Test-Path -LiteralPath $PartsPath)) { Remove-Item -LiteralPath $PartsPath -Recurse -Force -ErrorAction SilentlyContinue }
    [pscustomobject]@{ Directory = $Directory; Files = $files; MailboxFiles = $mailboxFiles; Preview = $preview }
}
