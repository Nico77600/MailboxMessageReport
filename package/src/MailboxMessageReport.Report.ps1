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
    then deleted. A HTML report holds every message, up to Report.HtmlMaxMessages (500,000 by default), in compressed
    blocks of columns (MailboxMessageReportNative.HtmlReport): the page shows them in a virtual table, with filters.
    CSV files: UTF-8 with BOM, configurable delimiter, text cells starting with = + - @ are prefixed with an apostrophe
    (no formula injection when opened in Excel), a cell is never longer than an Excel cell (a list of 10,000 recipients
    is cut with the number of the others; RecipientCount gives them all). No token or secret is ever part of the result.

.NOTES
    Author  : Nicolas Fabert
    Version : 2.1.1
#>

$script:ReportColumns = [ordered]@{
    Mailboxes = @('Input', 'Address', 'DisplayName', 'Status', 'State', 'Detail', 'Archive', 'ArchiveSource', 'ArchiveGuid', 'PrimaryFolders', 'ArchiveFolders', 'FoldersRead', 'FoldersFailed', 'PrimaryMessages', 'ArchiveMessages', 'RecoverableMessages', 'Messages', 'Notes')
    Folders   = @('Mailbox', 'MailboxName', 'Location', 'RecoverableItems', 'Path', 'Name', 'TotalItems', 'Messages', 'Seconds', 'Status', 'Detail')
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

# Messages per compressed block of a HTML report (the page decompresses them one after the other, with a progress bar).
$script:HtmlBlockRows = 20000

function Open-MmrHtmlReport {
    <#
        A HTML report opened for writing: the template up to the messages (its title), the messages added while the part
        files are merged (compiled: HtmlReport.Add), closed by Close-MmrHtmlReport.
    #>
    param([Parameter(Mandatory = $true)][string]$Path, [Parameter(Mandatory = $true)][string]$Title, [int]$Max)
    [MailboxMessageReportNative.HtmlReport]::new((Join-Path $script:ToolRoot 'templates\Report.template.html'), $Path, $Max, $script:HtmlBlockRows, [string[]]@('{{TITLE}}'), [string[]]@([Net.WebUtility]::HtmlEncode($Title)))
}

function Close-MmrHtmlReport {
    <# The end of a HTML report: the last block of messages, then the summary, the mailboxes and the folders. #>
    param(
        [Parameter(Mandatory = $true)][MailboxMessageReportNative.HtmlReport]$Report,
        [Parameter(Mandatory = $true)][System.Collections.IDictionary]$Summary,
        [AllowEmptyCollection()][object[]]$Mailboxes,
        [AllowEmptyCollection()][object[]]$Folders
    )
    $mbx = @($Mailboxes | Select-Object -Property ($script:ReportColumns.Mailboxes + 'Html', 'Csv'))
    $fld = @($Folders | Select-Object -Property $script:ReportColumns.Folders)
    $Report.Close([string[]]@('{{SUMMARY_JSON}}', '{{MAILBOXES_JSON}}', '{{FOLDERS_JSON}}'), [string[]]@((ConvertTo-MmrEmbeddedJson $Summary), (ConvertTo-MmrEmbeddedJson $mbx), (ConvertTo-MmrEmbeddedJson $fld)))
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
    .PARAMETER HtmlMaxMessages
        The most messages of a HTML report (every message below it); the CSV files hold them all.
    .PARAMETER PreviewMessages
        How many messages to keep for the window at most.
    .PARAMETER PreviewPerFolder
        For the window: the first messages of each folder (newest first), up to PreviewMessages in all. 0: the first
        PreviewMessages messages of the report.
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
        [int]$HtmlMaxMessages = 500000,
        [int]$PreviewMessages = 0,
        [int]$PreviewPerFolder = 0,
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
    $globalHtml = $null
    $htmlPath = Join-Path $Directory "$Prefix.html"
    $names = @($Result.Mailboxes | ForEach-Object { if ($_.DisplayName) { $_.DisplayName } else { $_.Address } })
    $globalTitle = "Mailbox Message Report | $(if ($names.Count -le 3) { $names -join ', ' } else { "$($names.Count) mailboxes" })"
    $preview = [MailboxMessageReportNative.RowBuffer]::new([Math]::Max(0, $PreviewMessages))
    try {
        if ($global -and $csv) { $files.Messages = Join-Path $Directory "$Prefix-Messages.csv"; $globalCsv = [MailboxMessageReportNative.CsvTarget]::new($files.Messages, [MailboxMessageReportNative.Columns]::Messages, $Delimiter) }
        if ($global -and $html) { $globalHtml = Open-MmrHtmlReport -Path $htmlPath -Title $globalTitle -Max $HtmlMaxMessages }
        $sub = Join-Path $Directory 'Mailboxes'
        if ($perMailbox) { [void][IO.Directory]::CreateDirectory($sub) }
        $done = 0
        foreach ($m in @($Result.Mailboxes)) {
            $done++
            Write-MmrProgress ($done / [Math]::Max(1, @($Result.Mailboxes).Count)) ('{0:N0}/{1:N0} mailboxes written' -f $done, @($Result.Mailboxes).Count)
            # A mailbox not read has no report of its own: the summary says why.
            $own = $perMailbox -and $m.State -eq 'Ok'
            $mailboxCsv = $null
            $mailboxHtml = $null
            $name = Get-MmrSafeFileName $m.Address
            $folders = if ($byMailbox.ContainsKey($m.Address)) { $byMailbox[$m.Address] } else { @() }
            try {
                if ($own -and $csv) { $m.Csv = "Mailboxes/$Prefix-$name.csv"; $mailboxCsv = [MailboxMessageReportNative.CsvTarget]::new((Join-Path $sub "$Prefix-$name.csv"), [MailboxMessageReportNative.Columns]::Messages, $Delimiter) }
                if ($own -and $html) { $mailboxHtml = Open-MmrHtmlReport -Path (Join-Path $sub "$Prefix-$name.html") -Title "Mailbox Message Report | $(if ($m.DisplayName) { $m.DisplayName } else { $m.Address })" -Max $HtmlMaxMessages }
                $targets = [MailboxMessageReportNative.CsvTarget[]]@(@($globalCsv, $mailboxCsv) | Where-Object { $_ })
                $reports = [MailboxMessageReportNative.HtmlReport[]]@(@($globalHtml, $mailboxHtml) | Where-Object { $_ })
                $buffers = [MailboxMessageReportNative.RowBuffer[]]@(@($(if ($PreviewPerFolder -le 0) { $preview })) | Where-Object { $_ })
                foreach ($f in $folders) {
                    if ($f.Messages -le 0) { continue }
                    # The window: the first messages of each folder (its sample joins the preview while there is room).
                    $sample = if ($PreviewPerFolder -gt 0 -and $preview.Lines.Count -lt $preview.Max) { [MailboxMessageReportNative.RowBuffer]::new($PreviewPerFolder) } else { $null }
                    $withSample = if ($sample) { [MailboxMessageReportNative.RowBuffer[]]@(@($buffers) + $sample) } else { $buffers }
                    foreach ($part in @($f.Parts)) { [void][MailboxMessageReportNative.Merge]::AppendPart($part, $targets, $withSample, $reports) }
                    if ($sample) { $preview.AddFrom($sample) }
                }
                if ($mailboxHtml) {
                    $m.Html = "Mailboxes/$Prefix-$name.html"
                    $counts = Get-MmrCounts -Mailboxes @($m) -Folders @($folders)
                    $summary = Get-MmrHtmlSummary -Result $Result -Mailbox $m -Counts $counts -Shown $mailboxHtml.Rows -Total $mailboxHtml.Seen -Layout $Layout -MessagesHere $true
                    Close-MmrHtmlReport -Report $mailboxHtml -Summary $summary -Mailboxes @($m) -Folders @($folders)
                    if (-not $csv) { $mailboxFiles++ }
                }
            }
            finally {
                if ($mailboxCsv) { $mailboxCsv.Dispose(); $mailboxFiles++ }
                if ($mailboxHtml) { $mailboxHtml.Dispose() }
            }
        }
    }
    catch {
        if ($globalHtml) { $globalHtml.Dispose() }
        throw
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
        # PerMailbox: the summary of the run, without its messages (they are in the report of each mailbox).
        if (-not $globalHtml) { $globalHtml = Open-MmrHtmlReport -Path $htmlPath -Title $globalTitle -Max 0 }
        $files.Html = $htmlPath
        $shown = if ($global) { $globalHtml.Rows } else { 0 }
        $total = if ($global) { $globalHtml.Seen } else { [long]$Result.Counts.Messages }
        $summary = Get-MmrHtmlSummary -Result $Result -Counts $Result.Counts -Shown $shown -Total $total -Layout $Layout -MessagesHere $global
        try { Close-MmrHtmlReport -Report $globalHtml -Summary $summary -Mailboxes @($Result.Mailboxes) -Folders @($Result.Folders) }
        finally { $globalHtml.Dispose() }
    }
    $files.Summary = Join-Path $Directory "$Prefix-Summary.json"
    $Result.RunFolder = $Directory
    $Result.Files = $files
    $data = [ordered]@{}
    foreach ($p in $Result.PSObject.Properties) { $data[$p.Name] = $p.Value }
    # The part files are gone once the report is written.
    $data.Folders = @($Result.Folders | Select-Object -Property * -ExcludeProperty Parts, SlicesDone, Started)
    [IO.File]::WriteAllText($files.Summary, (ConvertTo-Json -InputObject $data -Depth 8), [Text.UTF8Encoding]::new($false))

    if ($PartsPath -and (Test-Path -LiteralPath $PartsPath)) { Remove-Item -LiteralPath $PartsPath -Recurse -Force -ErrorAction SilentlyContinue }
    [pscustomobject]@{ Directory = $Directory; Files = $files; MailboxFiles = $mailboxFiles; Preview = $preview }
}

# SIG # Begin signature block
# MIImIwYJKoZIhvcNAQcCoIImFDCCJhACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCGq0INIdghEkyq
# 9mNXZp3A2/RQsJHoRypXB5Be9EYC6qCCDW8wggZyMIIEWqADAgECAghkM1HTxzif
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
# wZislcdCZuS6MYIYCjCCGAYCAQEwgYwweDELMAkGA1UEBhMCVVMxDjAMBgNVBAgM
# BVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMREwDwYDVQQKDAhTU0wgQ29ycDE0MDIG
# A1UEAwwrU1NMLmNvbSBDb2RlIFNpZ25pbmcgSW50ZXJtZWRpYXRlIENBIFJTQSBS
# MQIQBof4ljXihJpmCqqt2mhAWzANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3
# AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisG
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCBlDYe4BLT
# fvoJvQZQoj7Nq5476FCqyHuHAah+ippstDANBgkqhkiG9w0BAQEFAASCAgBoxuAH
# 8P6P8CIIFa0WQt8ixzz6bUQFpTitnF5jovjuMhnwW+LjYx3Rne3uqXRMIte6qWgT
# 5YvTbtbL255TDWGrKec1kRN804WhbTKbeBFuFXf4nFS2KX9TfxOJS9yKZD+d+HMo
# IKBkYIZveBH6JYmye8PbqAMh095BcvsgTjDXtBvX2w7PQyC9FeGjt4CdhjWoKOhA
# JDxRCSWysbR2Eqrg98S/bBodqyB/UtOpLGa9CskJc0tM2y2apKsKEECpNpySlaP/
# bM+a8fHA4I5Wr0LcENYdRBlVsku4vHxJK7dqp/5amJebX1DImyFyDtIzZDpCHrsi
# BD4bciir1jhJ2r6Nh0q2CxA+RAqqcyGLJHLCxIz5L24srrhDXKYPsFMpslEJIKBz
# 8KonUT8nmgzFWj5/9gh/cKCvzSGqTYXXbsO+gc9BkPFSuPg9T4I08Cj1OFcrlLNg
# qtNQmv1JBv2xSW8qRtstxRv0GkSyW4VOlHKIBoj5jr+E3p8S64nuVM9sy0rL4h/3
# 1ymTsEGRMqbgwttSW0Q+jRqjUNg3lnAA/LuK5Cee1pOmWJUJYa3H4nDLwhIZ7+8s
# ZY4U4goVMvrjgfjKrEdvVS4CfmzLoNGh1fU5i6iQLvdRvFODLWBuTbJgdA3S8l/Q
# 2Ga/eQ2fVjPzzJ3zbZuWtQhRdUUe/vfZeZ75qaGCFMcwghTDBgorBgEEAYI3AwMB
# MYIUszCCFK8GCSqGSIb3DQEHAqCCFKAwghScAgEDMQ0wCwYJYIZIAWUDBAIBMIGC
# BgsqhkiG9w0BCRABBKBzBHEwbwIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEIG27f5CCpMitBUcKsdN5lUqAg0JC8a+jK0k407EwwlFKAghZnYGb+BTx
# bBgPMjAyNjExMDUxNzUxMjNaMAMCAQECCQDianyVnpDNjaCCEacwggTzMIIC26AD
# AgECAhBUQ/g3vVruGAGEcrhc2o7jMA0GCSqGSIb3DQEBCwUAMHIxCzAJBgNVBAYT
# AlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91c3RvbjERMA8GA1UECgwI
# U1NMIENvcnAxLjAsBgNVBAMMJVNTTC5jb20gVGltZXN0YW1waW5nIElzc3Vpbmcg
# UlNBIENBIDIwHhcNMjYwNjEyMTcyMjQ2WhcNMzYwNjA5MTcyMjQ1WjBuMQswCQYD
# VQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNVBAcMB0hvdXN0b24xETAPBgNV
# BAoMCFNTTCBDb3JwMSowKAYDVQQDDCFTU0wuY29tIFRpbWVzdGFtcGluZyBVbml0
# IDIwMjYgRTEwWTATBgcqhkjOPQIBBggqhkjOPQMBBwNCAAToJYu4ZhMuSJI48ws2
# yQ4G2FewiFFIDjp945DC9ItwNo2Hq/FTMEoTHUp0UFEWzs4eucf8rRUDoxsNAgWl
# SDQ9o4IBUjCCAU4wHwYDVR0jBBgwFoAUcJplwUOAo/zpsWSJbvEwnAkhyXwwTQYI
# KwYBBQUHAQEEQTA/MD0GCCsGAQUFBzAChjFodHRwOi8vY2VydC5zc2wuY29tL1NT
# TC5jb20tdGltZVN0YW1waW5nLUktUjIuY2VyMFEGA1UdIARKMEgwPAYMKwYBBAGC
# qTABAwYBMCwwKgYIKwYBBQUHAgEWHmh0dHBzOi8vd3d3LnNzbC5jb20vcmVwb3Np
# dG9yeTAIBgZngQwBBAIwFgYDVR0lAQH/BAwwCgYIKwYBBQUHAwgwQgYDVR0fBDsw
# OTA3oDWgM4YxaHR0cDovL2NybHMuc3NsLmNvbS9TU0wuY29tLXRpbWVTdGFtcGlu
# Zy1JLVIyLmNybDAdBgNVHQ4EFgQUJ6wDVrooaroXO4M9QH95qCTswzwwDgYDVR0P
# AQH/BAQDAgeAMA0GCSqGSIb3DQEBCwUAA4ICAQBPdm7O/AOewhlNcGf6GOa4FeDM
# HXZ9tjFP25h7pm+G+ltlYpzINCBpwROQ/DyTrJ6qdgJjIVnTreUw8SoCEAPacwbr
# XhmHs+0db0Nap+yCK/Fs9gF4363tTGCRhPd7GltYx08SoG1ZXIc68twq8zGAlZ8m
# MA3BWTBUN+TL+Ptx/bz6KKY3w8pztL+7z8t6dFGdXZbZh+dxGeTP0J+zmdEPNpmC
# 6IeuunYbHCCc1BIfgYGMc6Xc/MuydVsahkqLhkSOBe6mziYYY4rv4YHYcCVCf45L
# 3jhweqWQfSASW3ZZ7OgI1GkQumH2prG4GeNISpdZZ+ATrIYHA7S7Eo2qlZVMeYDC
# vpK8E+h/G5tci4Sf2K9HgIs64DFRKsnjVkGD+qobSgmMXoaREpvGrtzJZKQRlgmG
# YoKDYeDl9cdqA7GLBuafkk9nW+FYKtuTe/91PHaC9Mrn56sLkf3+tNyWNQjvvrJ3
# at+GN+Sm/ICxu1ezgPueufDyL3Civn5921FIkFXGrMubKQV+2gvRSnZD763IoJR2
# uJt1yesHgcpLez63cCN1xsDToMIvwaJ9CL18vTJEFE2UDijoE/5KvvnkxoI2NRFh
# kO8PLEkcEFPtyqNkjCyJHyz7rzrSegVsEG12B7YjDsI8L/VDfUjbuFWiVLsHTJqj
# e/xfYy1bLeaxmVdNwTCCBd0wggPFoAMCAQICCHssm9MWgDKZMA0GCSqGSIb3DQEB
# CwUAMHwxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91
# c3RvbjEYMBYGA1UECgwPU1NMIENvcnBvcmF0aW9uMTEwLwYDVQQDDChTU0wuY29t
# IFJvb3QgQ2VydGlmaWNhdGlvbiBBdXRob3JpdHkgUlNBMB4XDTE2MDIxMjE3Mzkz
# OVoXDTQxMDIxMjE3MzkzOVowfDELMAkGA1UEBhMCVVMxDjAMBgNVBAgMBVRleGFz
# MRAwDgYDVQQHDAdIb3VzdG9uMRgwFgYDVQQKDA9TU0wgQ29ycG9yYXRpb24xMTAv
# BgNVBAMMKFNTTC5jb20gUm9vdCBDZXJ0aWZpY2F0aW9uIEF1dGhvcml0eSBSU0Ew
# ggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQD5D92jK33L0Cr+7GeFpucu
# G7p34eP1r6Ts+kpdkcRXR2sYd2t28v2T5D0PwhaeC2bDVpSeF4OFzlbv8hb9AGL1
# IglU6GUXTkG54E9Gl6obyLhuYl5psV/bKgJ+/GzK80HY7dDo/D9hSO2wAxQdEA5L
# GeC7TuyGZf82815nAgudhlVh/Xo47f7iGQC3b6FQYnV0PKD6yCWStG56Isf4HqHj
# st2RMasrHQT/pUoEN+mFpDMr/eLWVTR8GaRKaMeyqNO3yqGTiOvBl7yM+R3ZIoQk
# dMcEPWqpKZPM67hb4f5fJao0WMjBI1SdG5gRwzicfj2GbKUPQIZ8AvRcAk8oy65x
# nw86yDP+ESU16vy6xWA92XwY1bKp03V4A3IiyjrDH+8s5S6p+p4stlFG/a8D1upg
# aOqFFjZrhekewLPdxCTcgCqBQW2UPsjgyYFBAJ5ev3/FCJiiGCxCQLP5bzgnS06A
# 9D2BR+CIfOoczrV1XFEuHCt/GnIo5wC10XTG1+SfrQeTtlM1Nfw35MP2XRa+IXPe
# kgr4oGNqvJaSaj74vGVVm971DYkmBPwlGqYlacvCbcp84llfl6zr7y7IvNcbWTwr
# zPIZyJNrJ2MZz/zpJvjKcZt/k/40Z4ROmev8s3gJM3C6ZqZ27Rtz6xqlDcQiEyCU
# VgpOLGxOsf3PnAm6ojPthwIDAQABo2MwYTAdBgNVHQ4EFgQU3QQJB6L1en1SUxKS
# le44gCUNplkwDwYDVR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBTdBAkHovV6fVJT
# EpKV7jiAJQ2mWTAOBgNVHQ8BAf8EBAMCAYYwDQYJKoZIhvcNAQELBQADggIBACAY
# EZQp+yadHB4ecGHxlXKTcSStaJNYjjKvG7NwA/wlK3SFkD14avS5i6WXO7UYkbse
# p/lAW5H5VZmvHhHQXB2nZuOxlAcMMjmm6huwedgdnHBE44rdxPmVH4o4Qz8BhaVH
# pz1GsrzlImj3e5zYLD4KIcgtM6y/xYGZMXTBdXHFvrHwI0X0nWv8GWOdo7wExhgL
# JbtTiQ+zgFDeRe5Ef6uUeGSY0/Yo3YfYcGV0+w65E+unD2GpMpbM3rvtY0wYu6lA
# 96BUbiCIcXUY6nq0NHLgIyd3XLaQ6oYlQKvvMw/Ln4K+oiD79rUtGubChbF0D/vI
# ZQKkUgFH3Ukiwb/Y62usft7sYzMVtyMIj8YPjUFa3Y7FuY/lRT9427rSG0Cx/nFN
# P+CBorpetOwV4JPdCB9+4VWZCyHek54K++ajSb02MP7nd7KgdZe1LYGIF2Ug99qQ
# AJ/JUswyyjV89T0P2CvX9SZsyQY0lhbqcFkaMnl5C7aIfw9SSD2/bNiiRC7RTrdy
# WNOJE5X+RKv414sbbpy8LKBb1WoAr1834dX6EAuYnIbnJo/O8OxuilcLgONOssCg
# Y2GQulVoN3RqtpLbn6GGIrZlJw7stp9CYORnwrXaQQvE04thG7z6H5Er10QHXrop
# rNnF6e9TSFrrgPEoWCHNsAZV+yc/U5BwqQQeVye5MIIGyzCCBLOgAwIBAgIQUp27
# wriJh/WDgOCw66IGNjANBgkqhkiG9w0BAQsFADB8MQswCQYDVQQGEwJVUzEOMAwG
# A1UECAwFVGV4YXMxEDAOBgNVBAcMB0hvdXN0b24xGDAWBgNVBAoMD1NTTCBDb3Jw
# b3JhdGlvbjExMC8GA1UEAwwoU1NMLmNvbSBSb290IENlcnRpZmljYXRpb24gQXV0
# aG9yaXR5IFJTQTAeFw0yNjA2MTIxNjQyMTVaFw00MTAyMTIxNzM5MzlaMHIxCzAJ
# BgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91c3RvbjERMA8G
# A1UECgwIU1NMIENvcnAxLjAsBgNVBAMMJVNTTC5jb20gVGltZXN0YW1waW5nIElz
# c3VpbmcgUlNBIENBIDIwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQCf
# L6vmc02QvDFGmzYHeVpQY1T3Lu9aInnTXfJ/SJfPrltB3Br1uOPyX9KR5eocap9K
# HYWhp77zQTyYMR/3egjCOzLYJfZxeYrGq0XwckZgHh/rJGEoQPnTCRs7KwvTYhd/
# 78RmdrjiN9+GFXPg5T+TnbFD8+ciD+4WFzZaT5Jz8IeoCv6VhWQ/PsvLlGoH1/VE
# oav0kLftbUuBV0cLnR00h2xrdBc+GW1hm5tHeekTgFt+LSOhRbmfk7RrLB43jlH9
# hUvEDRgaWnU4dV9WeyperxuGfDUtziSMnpc6BeU6XD9rzN7ntO9t25mXd6F7BOa8
# 7QCTL+eo5YNL0g8/F3F0bq+msFtK4eXE0A0EBsdOIWCR9mZ/R1liVP8CdkigXECr
# Qcgs+iV/g6BaARPvpKJZKWbZdhEhfBWOl6obSmW9fBEhAHNKIQmPtB5AWcK8fHiB
# Awg283f7Ib/WvGERtqMa97S1M/eVfg884dUATRAmnhnM8jVhgmSFVGuJVXFVFTgW
# yDgE/A/oj5guYdvW1nhEvA4U0uH+AlgOA+UKEKjE6iAXFeDA9lmsPIRFgHRp457F
# bQ4ppw3hFG71yoXYoq78BvkAvt5gMtUZQ+sQTJhsgwClUqql5tzT+MNCeIL7euRq
# 5ECOBVzybWKn8JCAbUPDhn26VVEBj1VHplrbouxNtQIDAQABo4IBUTCCAU0wEgYD
# VR0TAQH/BAgwBgEB/wIBADAfBgNVHSMEGDAWgBTdBAkHovV6fVJTEpKV7jiAJQ2m
# WTBNBggrBgEFBQcBAQRBMD8wPQYIKwYBBQUHMAKGMWh0dHA6Ly9jZXJ0LnNzbC5j
# b20vU1NMY29tLVJvb3RDQS1SU0EtNDA5Ni1SMS5jZXIwPwYDVR0gBDgwNjA0BgRV
# HSAAMCwwKgYIKwYBBQUHAgEWHmh0dHBzOi8vd3d3LnNzbC5jb20vcmVwb3NpdG9y
# eTATBgNVHSUEDDAKBggrBgEFBQcDCDBCBgNVHR8EOzA5MDegNaAzhjFodHRwOi8v
# Y3Jscy5zc2wuY29tL1NTTGNvbS1Sb290Q0EtUlNBLTQwOTYtUjEuY3JsMB0GA1Ud
# DgQWBBRwmmXBQ4Cj/OmxZIlu8TCcCSHJfDAOBgNVHQ8BAf8EBAMCAYYwDQYJKoZI
# hvcNAQELBQADggIBAGUUW80x6YDqIGIFijnogpCNaFa5Bpgq6jiql/Qok0REC3WD
# P4ewDZIvfED9X4QDxJE80/1K9fSKz7mYOymli+krJgapg/fpgVbSZsXUXRm1fG6Y
# JRQ79lXmYE0oN/lDg9jtBvR5YB6lRzSs0mv8y+flmJd8k/ENvbGZyDxpEPrnb+ev
# 46uTZQE/QTjeIl0jgsXtWzb1HdIZEthk7EzXkMjhB5uwpGXQoOV/g6ZgYIpzu6ZP
# GJuRm03ULUpaCd/2Np1T8nCfJJpgQGDHr+/4VULpIY43OJjtNy8j0xX6pjEdQ+U6
# boBpA/Z1LhuNZeebzqV0JtSTG04AK6ZJ9P3vTvsk2JVrs9nUgGR3/Y8TUALyCztT
# 79iTUE31dE8iLXmT+4DXEX2h557koaPKIYqSohPtbxCSBRPnN4KKcUqRwpp9kxxE
# 5JVile2a1CoYZHaDi49ionQrWGazqAlese3/finwz0+tXyg1Z7RmXUSsMy1ed1SD
# eqEkvikDdlSWjvDmUzmaX4DuKOv/caTa/hBGgN4TZKGZNd0eC2qjwv99PXX2ZhwR
# PJwTY1gZFewGXc3SvMOwdGvw/dOumuqFGtfTiJLQ/t5D6UPGBLfKEafc7NYaEO30
# Z/bHWp7T+p2EUBU6RfHMxnPhRq3rGwB/NfT7+bwYQpbk6PD8dj6Qu9XDNI87MYIC
# VjCCAlICAQEwgYYwcjELMAkGA1UEBhMCVVMxDjAMBgNVBAgMBVRleGFzMRAwDgYD
# VQQHDAdIb3VzdG9uMREwDwYDVQQKDAhTU0wgQ29ycDEuMCwGA1UEAwwlU1NMLmNv
# bSBUaW1lc3RhbXBpbmcgSXNzdWluZyBSU0EgQ0EgMgIQVEP4N71a7hgBhHK4XNqO
# 4zALBglghkgBZQMEAgGgggFgMBoGCSqGSIb3DQEJAzENBgsqhkiG9w0BCRABBDAc
# BgkqhkiG9w0BCQUxDxcNMjYxMTA1MTc1MTIzWjAoBgkqhkiG9w0BCTQxGzAZMAsG
# CWCGSAFlAwQCAaEKBggqhkjOPQQDAjAvBgkqhkiG9w0BCQQxIgQgajfFVZHP/5rl
# xt7YsXEYtAZ5XHU72I+jZ9QEh2T8+3gwgcgGCyqGSIb3DQEJEAIvMYG4MIG1MIGy
# MIGvBCBXT+oXqvpefGormwjJmLFzXkULOHPKROaiTKCTDmPzsDCBijB2pHQwcjEL
# MAkGA1UEBhMCVVMxDjAMBgNVBAgMBVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMREw
# DwYDVQQKDAhTU0wgQ29ycDEuMCwGA1UEAwwlU1NMLmNvbSBUaW1lc3RhbXBpbmcg
# SXNzdWluZyBSU0EgQ0EgMgIQVEP4N71a7hgBhHK4XNqO4zAKBggqhkjOPQQDAgRH
# MEUCIQDQXuPnjRBD0HlJM3e793z0uSSDVxOucj1k4OXqbj2AAAIgAN1oRv62qJk8
# 84RqCRwtsLRe91lH0CSxEDe/k3voYwg=
# SIG # End signature block
