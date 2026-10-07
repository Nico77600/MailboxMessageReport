<#
.SYNOPSIS
    Mailbox Message Report - folders and messages of the mailboxes (dot-sourced by MailboxMessageReport.psm1).

.DESCRIPTION
    How the messages are read, for each mailbox of the request:

      1. Folders   every mail folder of the primary mailbox and of the archive, at once: GET .../mailFolders/delta
                   (the whole tree, flat, with parentFolderId and totalItemCount, 200 folders a page), the path of each
                   folder rebuilt from its parents. Recoverable Items when asked: the folders under
                   mailFolders/recoverableitemsroot (Deletions, Purges, Versions, DiscoveryHolds, SubstrateHolds,
                   Calendar Logging), readable with Mail.ReadBasic.All (measured). A folder without items, or left out
                   by Search.ExcludeFolders, is not read.
      2. Messages  each folder, page after page (1,000 messages a page): GET .../mailFolders/{id}/messages with
                   $select (the columns of the report), $filter on the received date and the subjects, $orderby
                   receivedDateTime desc. Exchange accepts contains(subject) only after a condition on the date
                   (InefficientFilter otherwise): with subjects and no start date, the filter starts with
                   receivedDateTime ge 1900-01-01. Each page is written at once to the part file of its folder
                   (compiled: MailboxMessageReportNative.PartWriter); nothing grows in memory with the messages.

    16 requests in flight, 4 per mailbox (a primary mailbox and its archive are two mailboxes): Invoke-MmrGraphPaged.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.0.0
#>

$script:StepIndex = 0
$script:StepTotal = 0

function Initialize-MmrSteps { param([int]$Total) $script:StepIndex = 0; $script:StepTotal = $Total }

function Write-MmrNextStep {
    param([Parameter(Mandatory = $true)][string]$Title, [string]$Icon = 'Info')
    $script:StepIndex++
    Write-MmrStep -Number $script:StepIndex -Total ([Math]::Max($script:StepIndex, $script:StepTotal)) -Title $Title -Icon $Icon
}

function Get-MmrProperty {
    <# A property of an object or a dictionary, or $null (Set-StrictMode safe). #>
    param($Object, [string]$Name)
    if ($null -eq $Object) { return $null }
    if ($Object -is [System.Collections.IDictionary]) { if ($Object.Contains($Name)) { return $Object[$Name] } return $null }
    $p = $Object.PSObject.Properties[$Name]
    if ($p) { return $p.Value }
    return $null
}

function Get-MmrMessageFilter {
    <#
        The $filter of the messages (empty: every message): the received date first (Exchange refuses contains() before
        it), then the subjects, any of them. Dates in UTC.
    #>
    param([AllowNull()][object]$Start, [AllowNull()][object]$End, [string[]]$Subject)
    $parts = [Collections.Generic.List[string]]::new()
    $subjects = @($Subject | Where-Object { $_ })
    if ($null -ne $Start) { $parts.Add('receivedDateTime ge ' + ([datetime]$Start).ToString('yyyy-MM-ddTHH:mm:ssZ', [Globalization.CultureInfo]::InvariantCulture)) }
    elseif ($subjects.Count) { $parts.Add('receivedDateTime ge 1900-01-01T00:00:00Z') }
    if ($null -ne $End) { $parts.Add('receivedDateTime lt ' + ([datetime]$End).ToString('yyyy-MM-ddTHH:mm:ssZ', [Globalization.CultureInfo]::InvariantCulture)) }
    if ($subjects.Count) {
        $any = @($subjects | ForEach-Object { "contains(subject,'$($_.Replace("'", "''"))')" })
        $parts.Add($(if ($any.Count -gt 1) { '(' + ($any -join ' or ') + ')' } else { $any[0] }))
    }
    return ($parts -join ' and ')
}

function Get-MmrMessagesUrl {
    <#
        The first page of the messages of a folder. -NoRecipients: without To, Cc and Bcc (Exchange reads the recipients of
        each message for them: about 70 times slower on a folder of meeting messages, lab 2026-10-07).
    #>
    param([Parameter(Mandatory = $true)][string]$MailboxKey, [Parameter(Mandatory = $true)][string]$FolderId, [string]$Filter, [int]$PageSize = 250, [switch]$NoRecipients, [string]$Select, [string]$OrderBy = 'desc')
    if (-not $Select) {
        $Select = [MailboxMessageReportNative.Columns]::Select
        if ($NoRecipients) { $Select = (@($Select -split ',' | Where-Object { $_ -notin 'toRecipients', 'ccRecipients', 'bccRecipients' }) -join ',') }
    }
    $url = "$(Get-MmrUserPath $MailboxKey)/mailFolders/$([Uri]::EscapeDataString($FolderId))/messages?`$select=$Select&`$top=$PageSize&`$orderby=receivedDateTime%20$OrderBy"
    if ($Filter) { $url += "&`$filter=$([Uri]::EscapeDataString($Filter))" }
    return $url
}
function Test-MmrFolderExcluded {
    <# The path of a folder matches a pattern of Search.ExcludeFolders (wildcards, any case). #>
    param([string]$Path, [string[]]$Patterns)
    foreach ($p in @($Patterns)) { if ($p -and $Path -like $p) { return $true } }
    return $false
}

function Get-MmrFolders {
    <#
    .SYNOPSIS
        Every folder to read of the mailboxes given: primary mailbox and archive (Locations), Recoverable Items when asked.
    .OUTPUTS
        One object per folder: Mailbox, MailboxName, MailboxKey, Location (Primary | Archive), RecoverableItems, Id, Name,
        Path (\Inbox\Projects), TotalItems, Messages, Status ('' = to read, Empty, Excluded), Detail, Parts, Slices, Pages.
        A mailbox that cannot be read gets its State and Detail (the folders of its archive alone may fail).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Mailboxes, [Parameter(Mandatory = $true)][pscustomobject]$Request)

    $folders = [Collections.Generic.List[object]]::new()
    $raw = @{}          # job id -> list of the folders answered by Graph
    $jobs = [Collections.Generic.List[object]]::new()
    $parts = @{}        # job id -> @{ Mailbox; Location; Recoverable; Key; Status; Code; Message }
    $select = 'displayName,parentFolderId,totalItemCount,childFolderCount'
    $n = 0
    foreach ($m in $Mailboxes) {
        if ($m.State -ne 'Ok') { continue }
        $locations = [Collections.Generic.List[object]]::new()
        if (@($Request.Location) -contains 'Primary') { $locations.Add(@('Primary', $m.Key)) }
        if (@($Request.Location) -contains 'Archive' -and $m.ArchiveKey) { $locations.Add(@('Archive', $m.ArchiveKey)) }
        foreach ($l in $locations) {
            $id = "f$n"; $n++
            $parts[$id] = @{ Mailbox = $m; Location = $l[0]; Recoverable = $false; Key = $l[1]; Status = 0; Code = ''; Message = '' }
            $raw[$id] = [Collections.Generic.List[object]]::new()
            $jobs.Add((New-MmrPagedJob -Id $id -Url "$(Get-MmrUserPath $l[1])/mailFolders/delta?`$select=$select" -Headers @{ Prefer = 'odata.maxpagesize=200' }))
            if ($Request.RecoverableItems) {
                $id = "f$n"; $n++
                $parts[$id] = @{ Mailbox = $m; Location = $l[0]; Recoverable = $true; Key = $l[1]; Status = 0; Code = ''; Message = '' }
                $raw[$id] = [Collections.Generic.List[object]]::new()
                $jobs.Add((New-MmrPagedJob -Id $id -Url "$(Get-MmrUserPath $l[1])/mailFolders/recoverableitemsroot/childFolders?`$select=$select&`$top=100"))
            }
        }
    }
    $total = $jobs.Count
    $done = 0
    $onPage = {
        param($Job, [string]$Content)
        $body = ConvertFrom-MmrJson $Content
        foreach ($v in @(Get-MmrProperty $body 'value')) { if ($null -ne $v) { $raw[$Job.Id].Add($v) } }
        return [string](Get-MmrProperty $body '@odata.nextLink')
    }
    $onDone = {
        param($Job, [int]$Status, [string]$Code, [string]$Message)
        $p = $parts[$Job.Id]
        $p.Status = $Status; $p.Code = $Code; $p.Message = $Message
        $script:MmrFolderJobsDone++
    }
    $script:MmrFolderJobsDone = 0
    $onProgress = { Write-MmrProgress ($script:MmrFolderJobsDone / [Math]::Max(1, $total)) ('{0:N0}/{1:N0} folder trees read' -f $script:MmrFolderJobsDone, $total) }
    Invoke-MmrGraphPaged -Jobs $jobs.ToArray() -OnPage $onPage -OnDone $onDone -OnProgress $onProgress

    # ---- Recoverable Items: the sub-folders of its folders, when there are (one level at a time) ----------------
    $deeper = [Collections.Generic.List[object]]::new()
    foreach ($id in @($parts.Keys | Where-Object { $parts[$_].Recoverable -and $parts[$_].Status -eq 200 })) {
        foreach ($f in @($raw[$id] | Where-Object { [int](Get-MmrProperty $_ 'childFolderCount') -gt 0 })) {
            $sub = "f$n"; $n++
            $parts[$sub] = @{ Mailbox = $parts[$id].Mailbox; Location = $parts[$id].Location; Recoverable = $true; Key = $parts[$id].Key; Status = 0; Code = ''; Message = ''; Child = $true }
            $raw[$sub] = [Collections.Generic.List[object]]::new()
            $deeper.Add((New-MmrPagedJob -Id $sub -Url "$(Get-MmrUserPath $parts[$id].Key)/mailFolders/$([Uri]::EscapeDataString([string]$f.id))/childFolders?`$select=$select&`$top=100"))
        }
    }
    while ($deeper.Count) {
        $round = $deeper.ToArray(); $deeper.Clear()
        Invoke-MmrGraphPaged -Jobs $round -OnPage $onPage -OnDone $onDone
        foreach ($j in $round) {
            if ($parts[$j.Id].Status -ne 200) { continue }
            foreach ($f in @($raw[$j.Id] | Where-Object { [int](Get-MmrProperty $_ 'childFolderCount') -gt 0 })) {
                $sub = "f$n"; $n++
                $parts[$sub] = @{ Mailbox = $parts[$j.Id].Mailbox; Location = $parts[$j.Id].Location; Recoverable = $true; Key = $parts[$j.Id].Key; Status = 0; Code = ''; Message = ''; Child = $true }
                $raw[$sub] = [Collections.Generic.List[object]]::new()
                $deeper.Add((New-MmrPagedJob -Id $sub -Url "$(Get-MmrUserPath $parts[$j.Id].Key)/mailFolders/$([Uri]::EscapeDataString([string]$f.id))/childFolders?`$select=$select&`$top=100"))
            }
        }
    }

    # ---- the folders, with their path -------------------------------------------------------------------------
    $byMailboxLocation = @{}
    foreach ($id in $parts.Keys) {
        $p = $parts[$id]
        $k = "$($p.Mailbox.Key)|$($p.Location)|$($p.Recoverable)"
        if (-not $byMailboxLocation.ContainsKey($k)) { $byMailboxLocation[$k] = @{ Part = $p; Raw = [Collections.Generic.List[object]]::new(); Failed = $null; ChildErrors = 0 } }
        $entry = $byMailboxLocation[$k]
        if ($p.Status -eq 200) { foreach ($v in $raw[$id]) { $entry.Raw.Add($v) } }
        elseif ($p.ContainsKey('Child')) { $entry.ChildErrors++ }
        elseif (-not $entry.Failed) { $entry.Failed = $p }
    }
    foreach ($k in $byMailboxLocation.Keys) {
        $entry = $byMailboxLocation[$k]; $p = $entry.Part; $m = $p.Mailbox
        if ($entry.Failed) {
            $fail = $entry.Failed
            $problem = Get-MmrMailboxProblem $fail.Status $fail.Code $fail.Message
            if ($p.Location -eq 'Primary' -and -not $p.Recoverable) {
                $m.State = Get-MmrMailboxState $fail.Status $fail.Code
                $m.Detail = $problem
            }
            elseif ($p.Location -eq 'Archive' -and -not $p.Recoverable) {
                $m.ArchiveState = 'Failed'
                $m.Notes = "archive not read: $problem"
            }
            else { $m.Notes = (@($m.Notes, "Recoverable Items ($($p.Location)) not read: $problem") | Where-Object { $_ }) -join " $([char]0x00B7) " }
            continue
        }
        if ($entry.ChildErrors) { $m.Notes = (@($m.Notes, "Recoverable Items ($($p.Location)): $($entry.ChildErrors) sub-folder list(s) not read") | Where-Object { $_ }) -join " $([char]0x00B7) " }
        $byId = @{}
        foreach ($f in $entry.Raw) { $byId[[string]$f.id] = $f }
        $pathOf = @{}
        $resolve = $null
        $resolve = {
            param([string]$Id, [int]$Depth)
            if ($pathOf.ContainsKey($Id)) { return $pathOf[$Id] }
            $f = $byId[$Id]
            $parent = [string](Get-MmrProperty $f 'parentFolderId')
            $name = [string](Get-MmrProperty $f 'displayName')
            $prefix = if ($p.Recoverable) { '\Recoverable Items' } else { '' }
            $path = if ($parent -and $byId.ContainsKey($parent) -and $Depth -lt 64) { (& $resolve $parent ($Depth + 1)) + '\' + $name } else { "$prefix\$name" }
            $pathOf[$Id] = $path
            return $path
        }
        foreach ($f in $entry.Raw) {
            $id = [string]$f.id
            $path = & $resolve $id 0
            $count = [long](Get-MmrProperty $f 'totalItemCount')
            $folder = [pscustomobject]@{
                Mailbox = $m.Address; MailboxName = $m.DisplayName; MailboxKey = $p.Key; Location = $p.Location; RecoverableItems = [bool]$p.Recoverable
                Id = $id; Name = [string](Get-MmrProperty $f 'displayName'); Path = $path; TotalItems = $count; Messages = [long]0
                Status = ''; Detail = ''; Parts = @(); Slices = 1; SlicesDone = 0; Pages = 0
            }
            if (Test-MmrFolderExcluded -Path $path -Patterns $Request.ExcludeFolders) { $folder.Status = 'Excluded'; $folder.Detail = 'left out (excluded folders)' }
            elseif ($count -eq 0) { $folder.Status = 'Empty'; $folder.Detail = 'no item' }
            $folders.Add($folder)
            if ($p.Location -eq 'Archive') { $m.ArchiveFolders++ } else { $m.PrimaryFolders++ }
        }
    }
    # A mailbox that cannot be read: its archive and Recoverable Items failed for the same reason (one note only).
    foreach ($m in $Mailboxes) { if ($m.State -ne 'Ok') { $m.Notes = ''; $m.ArchiveState = '' } }
    # The order of the report: mailbox (as in the request), primary mailbox then archive, Recoverable Items last, path.
    $order = @{}
    for ($i = 0; $i -lt $Mailboxes.Count; $i++) { $order[$Mailboxes[$i].Address] = $i }
    $sorted = @($folders | Sort-Object { $order[$_.Mailbox] }, { if ($_.Location -eq 'Primary') { 0 } else { 1 } }, { [int]$_.RecoverableItems }, { $_.Path.ToLowerInvariant() })
    return , $sorted
}

function Get-MmrFolderSlices {
    <#
        A large folder (more than Graph.SplitFolderItems items) is read in slices of its received dates, side by side:
        the pages of one list come one after the other, the slices of a folder 4 at a time. Its oldest and newest
        message (of the filter) are read first (one message each, its date only); the period between them is cut in
        equal parts (one per SplitFolderItems items, 16 at most). Returns folder ID -> list of filters, newest slice first;
        a folder whose dates could not be read is read in one list.
    #>
    param([object[]]$Folders, [pscustomobject]$Request, [hashtable]$Settings, [string]$Filter)
    $split = [int]$Settings.SplitFolderItems
    $slices = @{}
    $big = @($Folders | Where-Object { $split -gt 0 -and $_.TotalItems -gt $split })
    if (-not $big.Count) { return $slices }
    $edges = @{}
    $jobs = foreach ($f in $big) {
        foreach ($order in 'desc', 'asc') {
            New-MmrPagedJob -Id "$order|$($f.MailboxKey)|$($f.Id)" -Url (Get-MmrMessagesUrl -MailboxKey $f.MailboxKey -FolderId $f.Id -Filter $Filter -PageSize 1 -Select 'receivedDateTime' -OrderBy $order) -Tag $f
        }
    }
    $onPage = {
        param($Job, [string]$Content)
        $m = [regex]::Match($Content, '"receivedDateTime"\s*:\s*"([^"]+)"')
        if ($m.Success) { $edges[$Job.Id] = [datetime]::Parse($m.Groups[1].Value, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AdjustToUniversal -bor [Globalization.DateTimeStyles]::AssumeUniversal) }
        return ''
    }
    Invoke-MmrGraphPaged -Jobs @($jobs) -OnPage $onPage -OnDone { param($Job, $Status, $Code, $Message) if ($Status -ne 200) { Write-MmrLog 'WARN' "Dates of $($Job.Tag.Path): $Status $Code" } }
    foreach ($f in $big) {
        $newest = $edges["desc|$($f.MailboxKey)|$($f.Id)"]; $oldest = $edges["asc|$($f.MailboxKey)|$($f.Id)"]
        if ($null -eq $newest -or $null -eq $oldest -or $newest -le $oldest) { continue }
        $n = [int][Math]::Min(16, [Math]::Ceiling($f.TotalItems / [double]$split))
        # Whole seconds: the dates of the filter have no fraction, two slices meet on the same second (lt / ge).
        $from = $oldest
        $to = $newest.AddSeconds(1)
        $step = [Math]::Max([TimeSpan]::TicksPerSecond, [long](($to - $from).Ticks / $n))
        $cuts = [Collections.Generic.List[datetime]]::new()
        $cuts.Add($from)
        for ($k = 1; $k -lt $n; $k++) {
            $t = $from.AddTicks($step * $k)
            $t = $t.AddTicks(-($t.Ticks % [TimeSpan]::TicksPerSecond))
            if ($t -gt $cuts[-1] -and $t -lt $to) { $cuts.Add($t) }
        }
        $cuts.Add($to)
        $list = for ($k = $cuts.Count - 2; $k -ge 0; $k--) { Get-MmrMessageFilter -Start $cuts[$k] -End $cuts[$k + 1] -Subject $Request.Subject }
        $slices[$f.Id] = @($list)
    }
    return $slices
}

function Read-MmrMessages {
    <#
    .SYNOPSIS
        Reads the messages of the folders given (Status '' only), each page written to the part file of its folder (or
        of its slice: a large folder is read in slices of its dates, Get-MmrFolderSlices).
    .PARAMETER PartsPath
        Folder of the part files (one per folder or slice with messages, deleted once the report is written).
    .OUTPUTS
        Nothing: each folder gets Messages, Pages, Parts (its part files, newest first), Status (Read | Failed) and
        Detail, its mailbox the counts.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Folders,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Mailboxes,
        [Parameter(Mandatory = $true)][pscustomobject]$Request,
        [Parameter(Mandatory = $true)][hashtable]$Settings,
        [Parameter(Mandatory = $true)][string]$PartsPath
    )

    [void][IO.Directory]::CreateDirectory($PartsPath)
    $zone = Get-MmrTimeZone $Settings.TimeZone
    $filter = Get-MmrMessageFilter -Start $Request.Start -End $Request.End -Subject $Request.Subject
    $noRecipients = -not [bool](Get-MmrProperty $Request 'Recipients')
    $byAddress = @{}
    foreach ($m in $Mailboxes) { $byAddress[$m.Address] = $m }
    $toRead = @($Folders | Where-Object { $_.Status -eq '' })
    $slices = Get-MmrFolderSlices -Folders $toRead -Request $Request -Settings $Settings -Filter $filter
    if ($slices.Count) { Write-MmrItem Info ('{0:N0} large folder(s) read in {1:N0} slices of their dates, side by side' -f $slices.Count, (@($slices.Values | ForEach-Object { $_.Count }) | Measure-Object -Sum).Sum) -Icon Folder }
    $jobs = [Collections.Generic.List[object]]::new()
    $n = 0
    foreach ($f in $toRead) {
        $filters = @(if ($slices.ContainsKey($f.Id)) { $slices[$f.Id] } else { $filter })
        $f.Parts = [Collections.Generic.List[string]]::new()
        $f.Slices = $filters.Count
        $f.SlicesDone = 0
        foreach ($sliceFilter in $filters) {
            $part = Join-Path $PartsPath ('{0:D6}.jsonl' -f $n)
            $f.Parts.Add($part)
            $jobs.Add((New-MmrPagedJob -Id "m$n" -Url (Get-MmrMessagesUrl -MailboxKey $f.MailboxKey -FolderId $f.Id -Filter $sliceFilter -PageSize ([int]$Settings.PageSize) -NoRecipients:$noRecipients) -Tag ([pscustomobject]@{ Folder = $f; Part = $part })))
            $n++
        }
    }
    $state = @{ Writers = @{}; Items = [long]0; Done = [double]0; Folders = 0; Messages = [long]0; Shown = [datetime]::MinValue }
    foreach ($f in $toRead) { $state.Items += [Math]::Max(1, $f.TotalItems) }
    # The body of each page goes to the compiled writer as it was received (bytes): never through a PowerShell string.
    $onPage = {
        param($Job, $Content)
        $f = $Job.Tag.Folder
        $w = $state.Writers[$Job.Id]
        if (-not $w) {
            $w = [MailboxMessageReportNative.PartWriter]::new($Job.Tag.Part, $f.Mailbox, $f.MailboxName, $f.Location, $f.RecoverableItems, $f.Path, $f.Name, $zone)
            $state.Writers[$Job.Id] = $w
        }
        $page = $w.AddPage($Content)
        $f.Messages += $page.Rows
        $f.Pages++
        $state.Messages += $page.Rows
        return $page.NextLink
    }
    $onDone = {
        param($Job, [int]$Status, [string]$Code, [string]$Message)
        $f = $Job.Tag.Folder
        $w = $state.Writers[$Job.Id]
        if ($w) { $w.Dispose(); $state.Writers.Remove($Job.Id) }
        if ($Status -ne 200) {
            $f.Status = 'Failed'
            $f.Detail = Get-MmrMailboxProblem $Status $Code $Message
            # The content of an expanded folder lives in an auxiliary archive: Graph answers an error for it.
            if ($f.Location -eq 'Archive' -and $Status -ge 500) { $f.Detail += ' (an auto-expanding archive keeps some folders in an auxiliary archive: not supported)' }
        }
        $state.Done += [Math]::Max(1, $f.TotalItems) / [double]$f.Slices
        $f.SlicesDone++
        if ($f.SlicesDone -ge $f.Slices) {
            if ($f.Status -ne 'Failed') { $f.Status = 'Read'; $f.Detail = $(if ($f.Slices -gt 1) { "read in $($f.Slices) slices of its dates" } else { '' }) }
            $state.Folders++
        }
    }
    $onProgress = {
        $now = [datetime]::UtcNow
        if (($now - $state.Shown).TotalMilliseconds -lt 200) { return }
        $state.Shown = $now
        # Folders done by their items, plus the messages already read of the lists in course.
        $running = [long]0
        foreach ($id in $state.Writers.Keys) { $w = $state.Writers[$id]; $running += $w.Count }
        Write-MmrProgress ([Math]::Min(1.0, ($state.Done + $running) / [Math]::Max(1, $state.Items))) ('{0:N0}/{1:N0} folders read {2} {3:N0} messages' -f $state.Folders, $toRead.Count, $script:Dot, $state.Messages)
    }
    try { Invoke-MmrGraphPaged -Jobs $jobs.ToArray() -OnPage $onPage -OnDone $onDone -OnProgress $onProgress }
    finally {
        foreach ($w in @($state.Writers.Values)) { $w.Dispose() }
        $state.Writers.Clear()
    }    foreach ($f in $Folders) {
        $m = $byAddress[$f.Mailbox]
        if (-not $m) { continue }
        if ($f.Status -eq 'Read') {
            $m.FoldersRead++
            if ($f.RecoverableItems) { $m.RecoverableMessages += $f.Messages }
            elseif ($f.Location -eq 'Archive') { $m.ArchiveMessages += $f.Messages }
            else { $m.PrimaryMessages += $f.Messages }
            $m.Messages += $f.Messages
        }
        elseif ($f.Status -eq 'Failed') { $m.FoldersFailed++ }
    }
}

function Get-MmrCounts {
    <# The totals of some mailboxes and their folders (the run, or one mailbox for its own report). #>
    param([AllowEmptyCollection()][object[]]$Mailboxes, [AllowEmptyCollection()][object[]]$Folders)
    $mbx = @($Mailboxes)
    $folders = @($Folders)
    [ordered]@{
        Mailboxes           = $mbx.Count
        MailboxesRead       = @($mbx | Where-Object Status -eq 'Read').Count
        MailboxesPartial    = @($mbx | Where-Object Status -eq 'Partial').Count
        MailboxesNotRead    = @($mbx | Where-Object Status -eq 'Not read').Count
        WithArchive         = @($mbx | Where-Object Archive -eq 'Yes').Count
        Folders             = $folders.Count
        FoldersRead         = @($folders | Where-Object Status -eq 'Read').Count
        FoldersWithMessages = @($folders | Where-Object { $_.Messages -gt 0 }).Count
        FoldersEmpty        = @($folders | Where-Object Status -eq 'Empty').Count
        FoldersExcluded     = @($folders | Where-Object Status -eq 'Excluded').Count
        FoldersFailed       = @($folders | Where-Object Status -eq 'Failed').Count
        Messages            = [long](($mbx | Measure-Object -Property Messages -Sum).Sum)
        PrimaryMessages     = [long](($mbx | Measure-Object -Property PrimaryMessages -Sum).Sum)
        ArchiveMessages     = [long](($mbx | Measure-Object -Property ArchiveMessages -Sum).Sum)
        RecoverableMessages = [long](($mbx | Measure-Object -Property RecoverableMessages -Sum).Sum)
    }
}

function Update-MmrResultCounts {
    <# The totals of a result and the status of each mailbox and of the run. #>
    param([Parameter(Mandatory = $true)][pscustomobject]$Result)
    foreach ($m in @($Result.Mailboxes)) {
        $m.Status = if ($m.State -ne 'Ok') { 'Not read' } elseif ($m.FoldersFailed -or $m.ArchiveState -eq 'Failed') { 'Partial' } else { 'Read' }
    }
    $Result.Counts = Get-MmrCounts -Mailboxes @($Result.Mailboxes) -Folders @($Result.Folders)
    if (-not $Result.Error) {
        $Result.Status = if ($Result.Counts.MailboxesNotRead -or $Result.Counts.MailboxesPartial -or $Result.Counts.FoldersFailed -or @($Result.Warnings).Count) { 'Warning' } else { 'Completed' }
    }
}

function New-MmrResult {
    <# The result of a run, before the search: what was asked, by which application. #>
    param([Parameter(Mandatory = $true)][hashtable]$Settings, [Parameter(Mandatory = $true)][pscustomobject]$Request)
    $g = $script:Graph
    [pscustomobject]@{
        Tool = 'Mailbox Message Report'; Version = $script:ToolVersion; Status = 'Running'; Error = ''
        StartedUtc = [datetime]::UtcNow; CompletedUtc = $null; DurationSeconds = 0
        Request = [ordered]@{
            Mailboxes = @($Request.Mailboxes).Count; MailboxFile = $Request.MailboxFile
            Start = $Request.Start; End = $Request.End
            StartText = if ($null -ne $Request.Start) { Format-MmrDate $Request.Start $Settings.TimeZone } else { '' }
            EndText = if ($null -ne $Request.End) { Format-MmrDate $Request.End $Settings.TimeZone -PeriodEnd } else { '' }
            Subject = @($Request.Subject); Location = @($Request.Location); RecoverableItems = [bool]$Request.RecoverableItems; Recipients = [bool](Get-MmrProperty $Request 'Recipients')
            ExcludeFolders = @($Request.ExcludeFolders); Layout = $Request.Layout; Formats = @($Request.Formats)
            TimeZone = $(if ($Settings.TimeZone) { [string]$Settings.TimeZone } else { [TimeZoneInfo]::Local.Id })
            Filter = Get-MmrMessageFilter -Start $Request.Start -End $Request.End -Subject $Request.Subject
        }
        Tenant = if ($g) { [string]$g.TenantGuid } else { [string]$Settings.TenantId }; Organization = [string]$Settings.Organization
        AppId = [string]$Settings.AppId; AppName = if ($g) { [string]$g.AppName } else { '' }; Roles = if ($g) { @($g.Roles) } else { @() }
        Mailboxes = @(); Folders = @(); Warnings = @(); Counts = [ordered]@{}
        RunFolder = ''; Files = [ordered]@{}
    }
}

function Find-MmrMessages {
    <#
    .SYNOPSIS
        Reads the messages of the request: mailboxes, folders, then the messages into part files under PartsPath.
        Needs Connect-MmrGraph first. Steps 2 to 4 of a run.
    .OUTPUTS
        The result (New-MmrResult): Mailboxes, Folders, Warnings, Counts, Status.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][hashtable]$Settings, [Parameter(Mandatory = $true)][pscustomobject]$Request, [Parameter(Mandatory = $true)][string]$PartsPath)

    $g = $script:Graph
    $dot = $script:Dot
    $result = New-MmrResult -Settings $Settings -Request $Request
    $warnings = [Collections.Generic.List[string]]::new()

    # ---- 2. Mailboxes ------------------------------------------------------------------------------------------
    Write-MmrNextStep $(if (@($Request.Mailboxes).Count -gt 1) { "Mailboxes ($(@($Request.Mailboxes).Count))" } else { 'Mailbox' }) 'User'
    if (-not $g.CanFindUsers) { $warnings.Add('No User.Read.All: the addresses are read as typed (no alias), and the archive only from the ArchiveGuid of the list.'); Write-MmrItem Warn $warnings[-1] }
    elseif (-not $g.CanReadUsers) { $warnings.Add('User.ReadBasic.All only: the archive ID needs User.Read.All; the archive is read only from the ArchiveGuid of the list.'); Write-MmrItem Warn $warnings[-1] }
    $mailboxes = Resolve-MmrMailboxes -Entries @($Request.Mailboxes)
    $result.Mailboxes = $mailboxes
    if (@($Request.Location) -contains 'Archive') {
        $unknown = @($mailboxes | Where-Object { $_.State -eq 'Ok' -and $_.Archive -eq 'Unknown' })
        if ($unknown.Count) { $warnings.Add(('{0} mailbox(es) whose archive is not known: {1}.' -f $unknown.Count, (($unknown | Select-Object -First 3 | ForEach-Object Address) -join ', '))) }
    }
    if ($mailboxes.Count -le 10) {
        foreach ($m in $mailboxes) {
            $name = if ($m.DisplayName) { "$($m.DisplayName) <$($m.Address)>" } else { $m.Address }
            $archive = switch ($m.Archive) { 'Yes' { "archive$(if ($m.ArchiveSource -eq 'File') { ' (ArchiveGuid of the list)' })" } 'No' { 'no archive' } default { 'archive not known' } }
            if ($m.State -eq 'Ok' -and -not $m.UserId -and $g.CanFindUsers) { Write-MmrItem Warn "$name $dot not found in the directory: read as typed" -Icon User }
            elseif ($m.State -eq 'Ok') { Write-MmrItem $(if ($m.Archive -eq 'Unknown' -and @($Request.Location) -contains 'Archive') { 'Warn' } else { 'Ok' }) "$name $dot primary mailbox $dot $archive" -Icon $(if ($m.Archive -eq 'Yes') { 'Archive' } else { 'User' }) }
            else { Write-MmrItem Warn "$name $dot $($m.Detail)" -Icon User }
        }
    }
    else {
        foreach ($m in $mailboxes) { Write-MmrLog 'INFO' "Mailbox $($m.Address): $($m.State), archive $($m.Archive) ($($m.ArchiveSource)), $($m.Detail)" }
        Write-MmrItem Ok ('{0:N0} mailboxes {1} {2:N0} with an archive {1} {3:N0} not reachable' -f $mailboxes.Count, $dot, @($mailboxes | Where-Object Archive -eq 'Yes').Count, @($mailboxes | Where-Object State -ne 'Ok').Count) -Icon People
    }

    # ---- 3. Folders --------------------------------------------------------------------------------------------
    Write-MmrNextStep 'Folders' 'Folder'
    $folders = Get-MmrFolders -Mailboxes $mailboxes -Request $Request
    $result.Folders = $folders
    foreach ($m in @($mailboxes | Where-Object { $_.State -ne 'Ok' -and $_.Detail })) { Write-MmrLog 'WARN' "Mailbox $($m.Address) not read: $($m.Detail)" }
    $failedMailboxes = @($mailboxes | Where-Object State -ne 'Ok')
    foreach ($m in ($failedMailboxes | Select-Object -First 10)) { Write-MmrItem Warn "$($m.Address) $dot not read: $($m.Detail)" -Icon User }
    foreach ($m in @($mailboxes | Where-Object Notes | Select-Object -First 10)) { if ($m.ArchiveState -eq 'Failed' -or $m.Notes -match 'not read') { Write-MmrItem Warn "$($m.Address) $dot $($m.Notes)" -Icon Archive } }
    $count = { param([string]$Location, [string]$Status) @($folders | Where-Object { $_.Location -eq $Location -and $_.Status -eq $Status }).Count }
    $primary = @($folders | Where-Object Location -eq 'Primary'); $archive = @($folders | Where-Object Location -eq 'Archive')
    if (@($Request.Location) -contains 'Primary') { Write-MmrItem Ok ('Primary mailbox: {0:N0} folders {1} {2:N0} to read {1} {3:N0} empty{4}' -f $primary.Count, $dot, (& $count 'Primary' ''), (& $count 'Primary' 'Empty'), $(if (& $count 'Primary' 'Excluded') { " $dot $(& $count 'Primary' 'Excluded') left out" })) -Icon Folder }
    if (@($Request.Location) -contains 'Archive' -and -not @($mailboxes | Where-Object { $_.State -eq 'Ok' -and $_.Archive -eq 'Yes' }).Count) { Write-MmrItem Info $(if (@($mailboxes | Where-Object { $_.State -eq 'Ok' -and $_.Archive -eq 'Unknown' }).Count) { 'Archive: not known (User.Read.All, or the ArchiveGuid in the list)' } else { 'Archive: none (no mailbox read has an archive)' }) -Icon Archive }
    elseif (@($Request.Location) -contains 'Archive') { Write-MmrItem $(if ($archive.Count) { 'Ok' } else { 'Warn' }) ('Archive: {0:N0} folders {1} {2:N0} to read {1} {3:N0} empty{4}' -f $archive.Count, $dot, (& $count 'Archive' ''), (& $count 'Archive' 'Empty'), $(if (& $count 'Archive' 'Excluded') { " $dot $(& $count 'Archive' 'Excluded') left out" })) -Icon Archive }
    if ($Request.RecoverableItems) { Write-MmrItem Info ('Recoverable Items: {0:N0} folders, {1:N0} to read' -f @($folders | Where-Object RecoverableItems).Count, @($folders | Where-Object { $_.RecoverableItems -and $_.Status -eq '' }).Count) -Icon Folder }

    # ---- 4. Messages -------------------------------------------------------------------------------------------
    Write-MmrNextStep 'Messages' 'Mail'
    $filterText = @(
        if ($null -ne $Request.Start -or $null -ne $Request.End) { "received $(if ($null -ne $Request.Start) { "from $($result.Request.StartText)" }) $(if ($null -ne $Request.End) { "to $($result.Request.EndText)" })".Trim() -replace '\s+', ' ' }
        if (@($Request.Subject).Count) { "subject contains $((@($Request.Subject) | ForEach-Object { "'$_'" }) -join ' or ')" }
    ) -join " $dot "
    Write-MmrItem Info $(if ($filterText) { $filterText } else { 'every message, whatever its date' }) -Icon Filter
    Read-MmrMessages -Folders $folders -Mailboxes $mailboxes -Request $Request -Settings $Settings -PartsPath $PartsPath
    $result.Warnings = $warnings.ToArray()
    Update-MmrResultCounts $result
    $n = $result.Counts
    Write-MmrItem Ok ('{0:N0} messages {1} {2:N0} in primary mailboxes {1} {3:N0} in archives{4}' -f $n.Messages, $dot, $n.PrimaryMessages, $n.ArchiveMessages, $(if ($Request.RecoverableItems) { " $dot $('{0:N0}' -f $n.RecoverableMessages) in Recoverable Items" })) -Icon Mail
    if ($n.FoldersFailed) {
        Write-MmrItem Warn ('{0:N0} folder(s) not read: see the Folders tab of the report.' -f $n.FoldersFailed)
        foreach ($f in @($folders | Where-Object Status -eq 'Failed' | Select-Object -First 5)) { Write-MmrItem Warn "$($f.Mailbox) $($f.Location) $($f.Path) $dot $($f.Detail)" -Icon Folder }
    }
    $result.CompletedUtc = [datetime]::UtcNow
    $result.DurationSeconds = [Math]::Round(($result.CompletedUtc - $result.StartedUtc).TotalSeconds, 1)
    return $result
}
