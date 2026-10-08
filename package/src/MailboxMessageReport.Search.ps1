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
    Version : 2.1.0
#>

$script:StepIndex = 0
# The progress of the messages is shown every 200 ms at most (the tests: every page).
$script:MmrProgressMilliseconds = 200
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
        # A page cut short: asked again by Invoke-MmrGraphPaged (never a list of folders silently incomplete).
        if ((Get-MmrProperty (Get-MmrProperty $body 'error') 'code') -eq 'InvalidJson') { throw [IO.InvalidDataException]::new('the JSON of the page is not complete') }
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
    .SYNOPSIS
        Cuts each large folder (more than Graph.SplitFolderItems messages) into slices of the same number of messages,
        read side by side.
    .DESCRIPTION
        The pages of one list come one after the other; the slices of a folder are read 4 at a time (the limit of
        Exchange Online per mailbox). Three steps, each read 4 at a time per mailbox and 16 in all:
          1. its newest message with the number of messages of the filter ($count=true), and its oldest message;
          2. the received date of the message at 1/n, 2/n... of the folder ($top=1&$skip=k*N/n, newest first): the
             cuts, so that every slice holds about N/n messages (n = N / SplitFolderItems, 16 at most);
          3. a folder whose number is not known is cut in equal periods between its oldest and newest message.
        Equal periods alone left a folder of meeting messages almost whole in one slice (39,372 of 39,419 messages in
        3 days of 6 months, lab 2026-10-07): read page after page, one request at a time. Read-MmrMessages cuts a
        slice again while it is read when its mailbox has free slots.
    .OUTPUTS
        Folder ID -> @{ Count (messages of the filter, -1: not known); Oldest (date of the oldest message, or $null);
        Slices (newest first, each @{ Filter; Start; End; Lower }) }. Start / End: the bounds of the filter ($null:
        open, up to the period asked); Lower: the oldest date the slice can hold (to cut it again). A folder of one
        slice is not in the result, but its Count.
    #>
    param([object[]]$Folders, [pscustomobject]$Request, [hashtable]$Settings, [string]$Filter)
    $split = [int]$Settings.SplitFolderItems
    $result = @{}
    $big = @($Folders | Where-Object { $split -gt 0 -and $_.TotalItems -gt $split })
    if (-not $big.Count) { return $result }
    $parse = { param([string]$Text) [datetime]::Parse($Text, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AdjustToUniversal -bor [Globalization.DateTimeStyles]::AssumeUniversal) }
    $dates = @{}; $counts = @{}
    $onPage = {
        param($Job, [string]$Content)
        $m = [regex]::Match($Content, '"receivedDateTime"\s*:\s*"([^"]+)"')
        if ($m.Success) { $dates[$Job.Id] = & $parse $m.Groups[1].Value }
        $n = [regex]::Match($Content, '"@odata\.count"\s*:\s*(\d+)')
        if ($n.Success) { $counts[$Job.Id] = [long]$n.Groups[1].Value }
        return ''
    }
    $onDone = { param($Job, $Status, $Code, $Message) if ($Status -ne 200) { Write-MmrLog 'WARN' "Slices of $($Job.Tag.Path): $Status $Code $Message" } }
    # ---- 1. newest (with the number of messages) and oldest --------------------------------------------------
    $jobs = foreach ($f in $big) {
        $key = "$($f.MailboxKey)|$($f.Id)"
        New-MmrPagedJob -Id "desc|$key" -Url ((Get-MmrMessagesUrl -MailboxKey $f.MailboxKey -FolderId $f.Id -Filter $Filter -PageSize 1 -Select 'receivedDateTime' -OrderBy 'desc') + '&$count=true') -Tag $f
        New-MmrPagedJob -Id "asc|$key" -Url (Get-MmrMessagesUrl -MailboxKey $f.MailboxKey -FolderId $f.Id -Filter $Filter -PageSize 1 -Select 'receivedDateTime' -OrderBy 'asc') -Tag $f
    }
    Invoke-MmrGraphPaged -Jobs @($jobs) -OnPage $onPage -OnDone $onDone
    # ---- 2. the dates of the cuts: the message at k*N/n ----------------------------------------------------------
    $plan = @{}
    $jobs = [Collections.Generic.List[object]]::new()
    foreach ($f in $big) {
        $key = "$($f.MailboxKey)|$($f.Id)"
        $count = if ($counts.ContainsKey("desc|$key")) { $counts["desc|$key"] } else { -1 }
        $n = [int][Math]::Min(16, [Math]::Ceiling($(if ($count -ge 0) { $count } else { $f.TotalItems }) / [double]$split))
        $plan[$key] = @{ Count = $count; N = $n; Newest = $dates["desc|$key"]; Oldest = $dates["asc|$key"] }
        if ($count -lt 0 -or $n -lt 2) { continue }
        for ($k = 1; $k -lt $n; $k++) {
            $skip = [long][Math]::Floor($k * $count / [double]$n)
            $jobs.Add((New-MmrPagedJob -Id "cut|$key|$k" -Url ((Get-MmrMessagesUrl -MailboxKey $f.MailboxKey -FolderId $f.Id -Filter $Filter -PageSize 1 -Select 'receivedDateTime' -OrderBy 'desc') + "&`$skip=$skip") -Tag $f))
        }
    }
    if ($jobs.Count) { Invoke-MmrGraphPaged -Jobs $jobs.ToArray() -OnPage $onPage -OnDone $onDone }
    # ---- 3. the slices ---------------------------------------------------------------------------------------------
    foreach ($f in $big) {
        $key = "$($f.MailboxKey)|$($f.Id)"
        $p = $plan[$key]
        $entry = @{ Count = $p.Count; Oldest = $p.Oldest; Slices = @() }
        $result[$f.Id] = $entry
        if ($p.N -lt 2 -or $null -eq $p.Newest -or $null -eq $p.Oldest -or $p.Newest -le $p.Oldest) { continue }
        # Whole seconds: the dates of the filter have no fraction, two slices meet on the same second (lt / ge).
        $whole = { param([datetime]$d) $d.AddTicks(-($d.Ticks % [TimeSpan]::TicksPerSecond)) }
        $cuts = [Collections.Generic.List[datetime]]::new()
        if ($p.Count -ge 0) {
            # Newest first: the date of the message at k*N/n is the lower bound of slice k.
            for ($k = 1; $k -lt $p.N; $k++) {
                $d = $dates["cut|$key|$k"]
                if ($null -eq $d) { continue }
                $d = & $whole $d
                if ($d -gt $p.Oldest -and $d -le $p.Newest -and ($cuts.Count -eq 0 -or $d -lt $cuts[-1])) { $cuts.Add($d) }
            }
        }
        else {
            $from = $p.Oldest; $to = $p.Newest.AddSeconds(1)
            $step = [Math]::Max([TimeSpan]::TicksPerSecond, [long](($to - $from).Ticks / $p.N))
            for ($k = $p.N - 1; $k -ge 1; $k--) {
                $d = & $whole ($from.AddTicks($step * $k))
                if ($d -gt $from -and $d -lt $to -and ($cuts.Count -eq 0 -or $d -lt $cuts[-1])) { $cuts.Add($d) }
            }
        }
        if (-not $cuts.Count) { continue }
        # The newest and the oldest slice are open (up to the period asked): a message newer or older than the dates
        # read first (arrived since, or one of them left out by Graph) is still read.
        $list = [Collections.Generic.List[object]]::new()
        for ($k = 0; $k -le $cuts.Count; $k++) {
            $end = if ($k -eq 0) { $Request.End } else { $cuts[$k - 1] }
            $start = if ($k -eq $cuts.Count) { $Request.Start } else { $cuts[$k] }
            $lower = if ($k -eq $cuts.Count) { $(if ($null -ne $Request.Start -and [datetime]$Request.Start -gt $p.Oldest) { [datetime]$Request.Start } else { & $whole $p.Oldest }) } else { $cuts[$k] }
            $list.Add(@{ Filter = (Get-MmrMessageFilter -Start $start -End $end -Subject $Request.Subject); Start = $start; End = $end; Lower = $lower })
        }
        $entry.Slices = $list.ToArray()
    }
    return $result
}

function Read-MmrMessages {
    <#
    .SYNOPSIS
        Reads the messages of the folders given (Status '' only), each page written to the part file of its folder (or
        of its slice: a large folder is read in slices of the same number of messages, Get-MmrFolderSlices).
    .DESCRIPTION
        While the messages are read, a mailbox whose last lists are read page after page (free slots: Exchange allows
        4 requests at a time per mailbox) gets one of its slices cut in two (OnIdle): the slice stops at the middle of
        what is left (the floor of its writer), a new slice reads the older half. The progress is the share of the
        messages read, folder by folder (the number of messages of the filter when known: $count=true on the first
        page of each list).
    .PARAMETER PartsPath
        Folder of the part files (one per folder or slice with messages, deleted once the report is written).
    .OUTPUTS
        Nothing: each folder gets Messages, Pages, Parts (its part files, newest first), Slices, Seconds, Status
        (Read | Failed) and Detail, its mailbox the counts.
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
    $pageSize = [int]$Settings.PageSize
    $byAddress = @{}
    foreach ($m in $Mailboxes) { $byAddress[$m.Address] = $m }
    $toRead = @($Folders | Where-Object { $_.Status -eq '' })
    $started = [datetime]::UtcNow
    $plans = Get-MmrFolderSlices -Folders $toRead -Request $Request -Settings $Settings -Filter $filter
    $sliced = @($plans.Values | Where-Object { $_.Slices.Count -gt 1 })
    if ($sliced.Count) { Write-MmrItem Info ('{0:N0} large folder(s) read in {1:N0} slices of the same number of messages, side by side' -f $sliced.Count, (@($sliced | ForEach-Object { $_.Slices.Count }) | Measure-Object -Sum).Sum) -Icon Folder }
    $state = @{ Writers = @{}; Active = @{}; Expected = [double]0; Done = [double]0; Folders = 0; Messages = [long]0; Shown = [datetime]::MinValue; Parts = 0 }
    $newPart = { $part = Join-Path $PartsPath ('{0:D6}.jsonl' -f $state.Parts); $state.Parts++; $part }
    # A list of a folder (or of a slice): Lower is the oldest date it can hold ($null: not known, never cut again).
    $newJob = {
        param($Folder, [string]$SliceFilter, $Start, $Lower, [string]$Part, [switch]$Count)
        $url = Get-MmrMessagesUrl -MailboxKey $Folder.MailboxKey -FolderId $Folder.Id -Filter $SliceFilter -PageSize $pageSize -NoRecipients:$noRecipients
        if ($Count) { $url += '&$count=true' }
        $job = New-MmrPagedJob -Id "m$($state.Parts)" -Url $url -Tag ([pscustomobject]@{ Folder = $Folder; Part = $Part; Start = $Start; Lower = $Lower; Floor = $null; Last = $null; More = $true; Pages = 0 })
        if (-not $state.Active.ContainsKey($job.Mailbox)) { $state.Active[$job.Mailbox] = [Collections.Generic.List[object]]::new() }
        $state.Active[$job.Mailbox].Add($job)
        $job
    }
    $jobs = [Collections.Generic.List[object]]::new()
    foreach ($f in $toRead) {
        $plan = $plans[$f.Id]
        $f.Parts = [Collections.Generic.List[string]]::new()
        $f.SlicesDone = 0
        foreach ($p in 'Unreadable', 'Incomplete') { $f | Add-Member -NotePropertyName $p -NotePropertyValue 0 -Force }
        # Expected: the messages of the filter (-1: not known yet, TotalItems counts instead); Seconds: the time to read it.
        $f | Add-Member -NotePropertyName Expected -NotePropertyValue $(if ($plan -and $plan.Count -ge 0) { [long]$plan.Count } else { [long]-1 }) -Force
        $f | Add-Member -NotePropertyName Seconds -NotePropertyValue ([double]0) -Force
        $f | Add-Member -NotePropertyName Started -NotePropertyValue $null -Force
        $state.Expected += $(if ($f.Expected -ge 0) { $f.Expected } else { [Math]::Max(0, $f.TotalItems) })
        if ($plan -and $plan.Slices.Count -gt 1) {
            $f.Slices = $plan.Slices.Count
            foreach ($s in $plan.Slices) { $part = & $newPart; $f.Parts.Add($part); $jobs.Add((& $newJob $f $s.Filter $s.Start $s.Lower $part)) }
        }
        else {
            $f.Slices = 1
            # One list: its first page gives the number of messages of the filter (progress); cut again if large.
            $lower = if ($plan -and $plan.Oldest) { $(if ($null -ne $Request.Start -and [datetime]$Request.Start -gt $plan.Oldest) { [datetime]$Request.Start } else { $plan.Oldest.AddTicks(-($plan.Oldest.Ticks % [TimeSpan]::TicksPerSecond)) }) } else { $null }
            $part = & $newPart; $f.Parts.Add($part); $jobs.Add((& $newJob $f $filter $Request.Start $lower $part -Count:($f.Expected -lt 0)))
        }
    }
    # The share of the messages read: folder by folder, the messages read (at most those expected); a folder read counts
    # what it holds. Never more than 100 %, never stuck at 100 % before the end.
    $expected = { param($f) if ($f.Expected -ge 0) { [double]$f.Expected } else { [double][Math]::Max(0, $f.TotalItems) } }
    $setExpected = {
        param($f, [long]$Value)
        $before = & $expected $f
        $doneBefore = [Math]::Min([double]$f.Messages, $before)
        $f.Expected = $Value
        $after = & $expected $f
        $state.Expected += $after - $before
        $state.Done += [Math]::Min([double]$f.Messages, $after) - $doneBefore
    }
    # The body of each page goes to the compiled writer as it was received (bytes): never through a PowerShell string.
    $onPage = {
        param($Job, $Content)
        $tag = $Job.Tag
        $f = $tag.Folder
        $w = $state.Writers[$Job.Id]
        if (-not $w) {
            $w = [MailboxMessageReportNative.PartWriter]::new($tag.Part, $f.Mailbox, $f.MailboxName, $f.Location, $f.RecoverableItems, $f.Path, $f.Name, $zone)
            if ($tag.Floor) { $w.FloorUtc = $tag.Floor }
            $state.Writers[$Job.Id] = $w
        }
        $page = $w.AddPage($Content)
        if (-not $f.Started -or ($Job.Started -and $Job.Started -lt $f.Started)) { $f.Started = $Job.Started }
        $limit = & $expected $f
        $state.Done += [Math]::Min([double]($f.Messages + $page.Rows), $limit) - [Math]::Min([double]$f.Messages, $limit)
        $f.Messages += $page.Rows
        $f.Pages++
        $state.Messages += $page.Rows
        $tag.Pages++
        if ($w.LastUtc) { $tag.Last = $w.LastUtc }
        $tag.More = [bool]$page.NextLink
        if ($page.Total -ge 0 -and $f.Slices -eq 1 -and $tag.Pages -eq 1) { & $setExpected $f $page.Total }
        # $count was for the first page only.
        return ($page.NextLink -replace '([?&])(?:\$|%24)count=true&?', '$1' -replace '[?&]$', '')
    }
    $onDone = {
        param($Job, [int]$Status, [string]$Code, [string]$Message)
        $f = $Job.Tag.Folder
        $w = $state.Writers[$Job.Id]
        if ($w) { $w.Dispose(); $state.Writers.Remove($Job.Id) }
        if ($state.Active.ContainsKey($Job.Mailbox)) { [void]$state.Active[$Job.Mailbox].Remove($Job) }
        $f.Unreadable += $Job.Skipped
        $f.Incomplete += $Job.Partial
        if ($Status -ne 200) {
            $f.Status = 'Failed'
            $f.Detail = Get-MmrMailboxProblem $Status $Code $Message
            # The content of an expanded folder lives in an auxiliary archive: Graph answers an error for it.
            if ($f.Location -eq 'Archive' -and $Status -ge 500) { $f.Detail += ' (an auto-expanding archive keeps some folders in an auxiliary archive: not supported)' }
        }
        if (-not $f.Started -or ($Job.Started -and $Job.Started -lt $f.Started)) { $f.Started = $Job.Started }
        $f.SlicesDone++
        if ($f.SlicesDone -ge $f.Slices) {
            if ($f.Status -ne 'Failed') {
                $f.Status = 'Read'
                $f.Detail = @(
                    if ($f.Slices -gt 1) { "read in $($f.Slices) slices" }
                    if ($f.Unreadable) { "$($f.Unreadable) message(s) left out: Graph cannot return them" }
                    if ($f.Incomplete) { "$($f.Incomplete) message(s) without sender and recipients: Graph cannot return them" }
                ) -join " $($script:Dot) "
            }
            if ($f.Started) { $f.Seconds = [Math]::Round(([datetime]::UtcNow - $f.Started).TotalSeconds, 1) }
            # Read: it holds what was read.
            & $setExpected $f $f.Messages
            $state.Folders++
        }
    }
    # A mailbox with free slots: the list of it with the longest period left is cut in two at the middle of that period
    # (whole seconds); the list in course stops there (the floor of its writer), a new one reads the older half.
    $onIdle = {
        param([string]$Mailbox, [int]$Free)
        $active = $state.Active[$Mailbox]
        if (-not $active) { return }
        $best = $null; $span = [TimeSpan]::FromSeconds(2)
        foreach ($j in $active) {
            $t = $j.Tag
            if (-not $t.More -or -not $t.Last -or -not $t.Lower -or $j.Cuts -or $t.Folder.Slices -ge 64) { continue }
            if (($t.Last - $t.Lower) -gt $span) { $span = $t.Last - $t.Lower; $best = $j }
        }
        if (-not $best) { return }
        $t = $best.Tag
        $middle = $t.Lower.AddTicks([long]($span.Ticks / 2))
        $middle = $middle.AddTicks(-($middle.Ticks % [TimeSpan]::TicksPerSecond))
        if ($middle -le $t.Lower -or $middle -gt $t.Last) { return }
        $f = $t.Folder
        $part = & $newPart
        $f.Parts.Insert($f.Parts.IndexOf($t.Part) + 1, $part)
        $f.Slices++
        $new = & $newJob $f (Get-MmrMessageFilter -Start $t.Start -End $middle -Subject $Request.Subject) $t.Start $t.Lower $part
        # The list in course now holds [middle, its end): a next cut of it starts there.
        $t.Floor = $middle
        $t.Start = $middle
        $t.Lower = $middle
        if ($state.Writers.ContainsKey($best.Id)) { $state.Writers[$best.Id].FloorUtc = $middle }
        Write-MmrLog 'INFO' ("{0} {1} {2}: a slice cut at {3:yyyy-MM-dd HH:mm:ss} UTC (free request slots of the mailbox), {4} slices" -f $f.Mailbox, $f.Location, $f.Path, $middle, $f.Slices)
        $new
    }
    $onProgress = {
        $now = [datetime]::UtcNow
        if (($now - $state.Shown).TotalMilliseconds -lt $script:MmrProgressMilliseconds) { return }
        $state.Shown = $now
        $rate = $state.Messages / [Math]::Max(1.0, ($now - $started).TotalSeconds)
        Write-MmrProgress ([Math]::Min(1.0, $state.Done / [Math]::Max(1.0, $state.Expected))) ('{0:N0}/{1:N0} folders read {2} {3:N0} messages {2} {4:N0} a second' -f $state.Folders, $toRead.Count, $script:Dot, $state.Messages, $rate)
    }
    try { Invoke-MmrGraphPaged -Jobs $jobs.ToArray() -OnPage $onPage -OnDone $onDone -OnProgress $onProgress -OnIdle $onIdle }
    finally {
        foreach ($w in @($state.Writers.Values)) { $w.Dispose() }
        $state.Writers.Clear()
    }
    foreach ($f in $Folders) {
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
    # ---- the speed, for the log and the console ---------------------------------------------------------------------
    $seconds = ([datetime]::UtcNow - $started).TotalSeconds
    $read = @($toRead | Where-Object { $_.PSObject.Properties['Started'] -and $_.Started })
    Write-MmrLog 'INFO' ('Messages: {0:N0} in {1} ({2:N0} a second)' -f $state.Messages, (Format-MmrDuration $seconds), ($state.Messages / [Math]::Max(0.1, $seconds)))
    foreach ($g in @($read | Group-Object Mailbox | Sort-Object { ($_.Group | Measure-Object Messages -Sum).Sum } -Descending | Select-Object -First 10)) {
        $first = ($g.Group | Measure-Object -Property Started -Minimum).Minimum
        $last = ($g.Group | ForEach-Object { $_.Started.AddSeconds($_.Seconds) } | Measure-Object -Maximum).Maximum
        $span = [Math]::Max(0.1, ($last - $first).TotalSeconds)
        $count = ($g.Group | Measure-Object Messages -Sum).Sum
        Write-MmrLog 'INFO' ('Mailbox {0}: {1:N0} messages in {2} ({3:N0} a second), {4} folder(s)' -f $g.Name, $count, (Format-MmrDuration $span), ($count / $span), $g.Count)
    }
    foreach ($f in @($read | Where-Object Messages | Sort-Object Seconds -Descending | Select-Object -First 5)) {
        Write-MmrLog 'INFO' ('Slow folder {0} {1} {2}: {3:N0} messages in {4} ({5:N0} a second), {6} slice(s), {7} page(s)' -f $f.Mailbox, $f.Location, $f.Path, $f.Messages, (Format-MmrDuration $f.Seconds), ($f.Messages / [Math]::Max(0.1, $f.Seconds)), $f.Slices, $f.Pages)
    }
    if ($state.Messages) { Write-MmrItem Info ('{0:N0} messages read in {1}: {2:N0} a second' -f $state.Messages, (Format-MmrDuration $seconds), ($state.Messages / [Math]::Max(0.1, $seconds))) -Icon Clock }
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
    # Messages that Graph could not return, even one at a time: the run finishes with warnings.
    $lost = @($folders | Where-Object { $_.PSObject.Properties['Unreadable'] -and $_.Unreadable -gt 0 })
    $partial = @($folders | Where-Object { $_.PSObject.Properties['Incomplete'] -and $_.Incomplete -gt 0 })
    if ($lost.Count) {
        $warnings.Add(('{0:N0} message(s) left out: Graph cannot return them, even one at a time ({1}). See the Folders tab and the log.' -f ($lost | Measure-Object -Property Unreadable -Sum).Sum, (($lost | Select-Object -First 3 | ForEach-Object { "$($_.Mailbox) $($_.Location) $($_.Path)" }) -join ', ')))
        Write-MmrItem Warn $warnings[-1] -Icon Folder
    }
    if ($partial.Count) {
        $warnings.Add(('{0:N0} message(s) read without their sender and recipients: Graph cannot return them ({1}). See the Folders tab.' -f ($partial | Measure-Object -Property Incomplete -Sum).Sum, (($partial | Select-Object -First 3 | ForEach-Object { "$($_.Mailbox) $($_.Location) $($_.Path)" }) -join ', ')))
        Write-MmrItem Warn $warnings[-1] -Icon Folder
    }
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
