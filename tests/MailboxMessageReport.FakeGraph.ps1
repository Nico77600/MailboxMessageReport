<#
    Mailbox Message Report - simulated Exchange Online tenant behind Microsoft Graph, for the tests.
    Author  : Nicolas Fabert
    Version : 1.1.1

    Start-MmrGraphSend (the only function of the tool that touches the network for Graph) is replaced by a mock that
    answers from this tenant: users, their primary mailbox and archive, folders, messages, $batch (v1.0 and beta). It
    reproduces the behaviour measured in the lab (2026-10-07):
      - /beta/users/{id}/settings/exchange gives primaryMailboxId and inPlaceArchiveMailboxId (MBX:<guid>@<tenant>);
        404 MailboxNotEnabledForRESTAPI for a mailbox inactive or on-premises;
      - a mailbox is opened by its user ID, any of its addresses, or its MBX ID; the archive by its MBX ID only;
        404 ErrorInvalidUser for an address that is not a mailbox;
      - mailFolders/delta: every folder of the tree, flat (msgfolderroot excluded), Prefer odata.maxpagesize (10 by
        default), a deltaLink on the last page; recoverableitemsroot/childFolders: the folders of Recoverable Items;
      - messages: $filter receivedDateTime ge / lt and contains(subject,'...') (any case, or), $orderby
        receivedDateTime desc, $top / $skip and nextLink; contains() before the date, or with $orderby and no date:
        400 InefficientFilter, as Exchange answers;
      - a page of messages cut short (200, the JSON ends after "value":[), as Exchange Online sometimes answers: a
        message it cannot return (Add-FakeMessage -Cut Always, or Recipients: only when its sender or recipients are
        asked), pages larger than CutAbove (a page that takes too long), or once (CutOnce).
    Every request is recorded (Calls), with the most requests in flight at once per mailbox (Open, MaxOpen).
#>

function Reset-FakeTenant {
    param([string]$TenantId = '11111111-2222-3333-4444-555555555555')
    $script:Fake = @{
        Tenant    = $TenantId
        Users     = [Collections.Generic.List[object]]::new()
        Mailboxes = @{}            # key (user ID, address, MBX ID, lower case) -> mailbox @{ Key; Kind; Folders; Owner }
        Calls     = [Collections.Generic.List[object]]::new()
        Throttle  = @{}            # url pattern -> number of 429 still to answer
        Fail      = @{}            # url pattern -> @{ Status; Code }
        CutOnce   = @{}            # url pattern -> number of pages still to cut short
        CutAbove  = @{}            # url pattern -> largest $top answered whole
        Open      = @{}            # mailbox -> requests in flight
        MaxOpen   = @{}            # mailbox -> most requests in flight at once
        Next      = 0
    }
}

function New-FakeId { param([string]$Prefix = 'AAMk') $script:Fake.Next++; return '{0}{1:D8}=' -f $Prefix, $script:Fake.Next }

function New-FakeMailboxStore {
    param([string]$Key, [string]$Kind)
    $root = New-FakeId 'ROOT'
    $rir = New-FakeId 'RIR'
    @{ Key = $Key; Kind = $Kind; Root = $root; RecoverableRoot = $rir; Folders = [Collections.Generic.List[object]]::new() }
}

function Add-FakeUser {
    <#
        A user with a primary mailbox (and an archive with -Archive). -OnPremises: the user exists, its mailbox cannot be
        opened (MailboxNotEnabledForRESTAPI). Returns @{ Id; Address; PrimaryKey; ArchiveKey; ArchiveGuid }.
    #>
    param([Parameter(Mandatory)][string]$Address, [string]$Name, [string[]]$Aliases = @(), [switch]$Archive, [switch]$OnPremises)
    $a = $Address.ToLowerInvariant()
    $id = [guid]::NewGuid().ToString()
    $exchangeGuid = [guid]::NewGuid().ToString()
    $archiveGuid = if ($Archive) { [guid]::NewGuid().ToString() } else { '' }
    $primaryKey = "MBX:$exchangeGuid@$($script:Fake.Tenant)"
    $archiveKey = if ($Archive) { "MBX:$archiveGuid@$($script:Fake.Tenant)" } else { '' }
    $user = [pscustomobject]@{
        id = $id; displayName = $(if ($Name) { $Name } else { $a }); mail = $a; userPrincipalName = $a
        proxyAddresses = @("SMTP:$a") + @($Aliases | ForEach-Object { "smtp:$($_.ToLowerInvariant())" })
        PrimaryKey = $primaryKey; ArchiveKey = $archiveKey; OnPremises = [bool]$OnPremises
    }
    $script:Fake.Users.Add($user)
    if (-not $OnPremises) {
        $primary = New-FakeMailboxStore -Key $primaryKey -Kind 'Primary'
        foreach ($k in @($id, $a, $primaryKey.ToLowerInvariant()) + @($Aliases | ForEach-Object { $_.ToLowerInvariant() })) { $script:Fake.Mailboxes[$k] = $primary }
        if ($Archive) { $script:Fake.Mailboxes[$archiveKey.ToLowerInvariant()] = New-FakeMailboxStore -Key $archiveKey -Kind 'Archive' }
    }
    [pscustomobject]@{ Id = $id; Address = $a; PrimaryKey = $primaryKey; ArchiveKey = $archiveKey; ArchiveGuid = $archiveGuid }
}

function Get-FakeStore {
    param([Parameter(Mandatory)][string]$Address, [ValidateSet('Primary', 'Archive')][string]$Location = 'Primary')
    $user = $script:Fake.Users | Where-Object { $_.mail -eq $Address.ToLowerInvariant() } | Select-Object -First 1
    if (-not $user) { throw "No fake user $Address" }
    $key = if ($Location -eq 'Archive') { $user.ArchiveKey } else { $user.PrimaryKey }
    return $script:Fake.Mailboxes[$key.ToLowerInvariant()]
}

function Add-FakeFolder {
    <# A folder by path (\Inbox\Projects), created with its parents; -Recoverable: under Recoverable Items (\Deletions). Returns the folder. #>
    param([Parameter(Mandatory)][string]$Address, [ValidateSet('Primary', 'Archive')][string]$Location = 'Primary', [Parameter(Mandatory)][string]$Path, [switch]$Recoverable)
    $store = Get-FakeStore $Address $Location
    $parent = if ($Recoverable) { $store.RecoverableRoot } else { $store.Root }
    $folder = $null
    foreach ($name in ($Path.Trim('\') -split '\\')) {
        $folder = $store.Folders | Where-Object { $_.parentFolderId -eq $parent -and $_.displayName -eq $name } | Select-Object -First 1
        if (-not $folder) {
            $folder = [pscustomobject]@{ id = (New-FakeId 'FLD'); displayName = $name; parentFolderId = $parent; Recoverable = [bool]$Recoverable; Messages = [Collections.Generic.List[object]]::new() }
            $store.Folders.Add($folder)
        }
        $parent = $folder.id
    }
    return $folder
}

function Add-FakeMessage {
    <# A message in a folder (created when missing). Dates are UTC. #>
    param(
        [Parameter(Mandatory)][string]$Address, [ValidateSet('Primary', 'Archive')][string]$Location = 'Primary', [Parameter(Mandatory)][string]$Path, [switch]$Recoverable,
        [Parameter(Mandatory)][string]$Subject, [Parameter(Mandatory)][datetime]$Received, [string]$From = 'sender@fabrikam.test', [string]$FromName = 'Sender',
        [string[]]$To = @(), [string[]]$Cc = @(), [string[]]$Bcc = @(), [switch]$Attachments, [string]$Type, [string]$Body,
        [ValidateSet('', 'Always', 'Recipients')][string]$Cut = ''
    )
    $folder = Add-FakeFolder -Address $Address -Location $Location -Path $Path -Recoverable:$Recoverable
    $utc = [datetime]::SpecifyKind($Received, [DateTimeKind]::Utc)
    $recipients = { param([string[]]$List) @($List | Where-Object { $_ } | ForEach-Object { @{ emailAddress = @{ name = $_; address = $_ } } }) }
    $m = [ordered]@{
        id = (New-FakeId 'MSG'); subject = $Subject
        receivedDateTime = $utc.ToString('yyyy-MM-ddTHH:mm:ssZ'); sentDateTime = $utc.AddMinutes(-1).ToString('yyyy-MM-ddTHH:mm:ssZ')
        from = @{ emailAddress = @{ name = $FromName; address = $From } }; sender = @{ emailAddress = @{ name = $FromName; address = $From } }
        toRecipients = @(& $recipients $To); ccRecipients = @(& $recipients $Cc); bccRecipients = @(& $recipients $Bcc)
        internetMessageId = "<$([guid]::NewGuid().ToString('n'))@fabrikam.test>"; hasAttachments = [bool]$Attachments; importance = 'normal'; isRead = $true
        body = @{ contentType = 'html'; content = $(if ($Body) { $Body } else { "<p>$Subject</p>" }) }
    }
    if ($Type) { $m['@odata.type'] = $Type }
    # Never answered (not a property of Graph): a page holding this message is cut short.
    if ($Cut) { $m['Cut'] = $Cut }
    $folder.Messages.Add([pscustomobject]$m)
    return [pscustomobject]$m
}

function ConvertTo-FakeJson { param($Value) if ($null -eq $Value) { return '' } return ($Value | ConvertTo-Json -Depth 12 -Compress) }

function Get-FakeQuery {
    param([string]$Query)
    $q = @{}
    foreach ($pair in ($Query.TrimStart('?') -split '&')) {
        if (-not $pair) { continue }
        $k, $v = $pair -split '=', 2
        $q[[Uri]::UnescapeDataString($k)] = [Uri]::UnescapeDataString(([string]$v).Replace('+', ' '))
    }
    return $q
}

function Test-FakeFilter {
    <# A message against the $filter of the tool (the subset Exchange accepts, in the order it accepts). #>
    param($Message, [string]$Filter)
    if (-not $Filter) { return $true }
    $received = [datetime]::Parse($Message.receivedDateTime, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AdjustToUniversal)
    foreach ($m in [regex]::Matches($Filter, "receivedDateTime (ge|lt) (\S+?Z)")) {
        $d = [datetime]::Parse($m.Groups[2].Value, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AdjustToUniversal)
        if ($m.Groups[1].Value -eq 'ge' -and $received -lt $d) { return $false }
        if ($m.Groups[1].Value -eq 'lt' -and $received -ge $d) { return $false }
    }
    $subjects = @([regex]::Matches($Filter, "contains\(subject,'((?:[^']|'')*)'\)") | ForEach-Object { $_.Groups[1].Value.Replace("''", "'") })
    if ($subjects.Count -and -not @($subjects | Where-Object { ([string]$Message.subject).IndexOf($_, [StringComparison]::OrdinalIgnoreCase) -ge 0 }).Count) { return $false }
    return $true
}

function Invoke-FakeGraphRequest {
    <# One Graph request against the simulated tenant: @{ status; body; headers }. Url relative to /v1.0 or /beta, or absolute. #>
    param([string]$Method, [string]$Url, $Body, [hashtable]$Headers = @{}, [string]$Version = 'v1.0')

    $relative = $Url -replace '^https://graph\.microsoft\.com/(v1\.0|beta)', ''
    if ($Url -match '^https://graph\.microsoft\.com/beta') { $Version = 'beta' }
    $decoded = [Uri]::UnescapeDataString($relative)
    $script:Fake.Calls.Add([pscustomobject]@{ Method = $Method; Url = $decoded; Version = $Version })
    $path, $query = $relative -split '\?', 2
    $q = Get-FakeQuery $query
    $segments = @($path.Trim('/') -split '/' | ForEach-Object { [Uri]::UnescapeDataString($_) })
    foreach ($pattern in @($script:Fake.Throttle.Keys)) {
        if ($decoded -like $pattern -and $script:Fake.Throttle[$pattern] -gt 0) {
            $script:Fake.Throttle[$pattern]--
            return @{ status = 429; headers = @{ 'Retry-After' = '0.05' }; body = @{ error = @{ code = 'ApplicationThrottled'; message = 'Too many requests' } } }
        }
    }
    foreach ($pattern in @($script:Fake.Fail.Keys)) {
        $f = $script:Fake.Fail[$pattern]
        if ($decoded -like $pattern) { return @{ status = $f.Status; body = @{ error = @{ code = $f.Code; message = 'simulated failure' } } } }
    }
    $error404 = { param([string]$Code = 'ErrorInvalidUser', [string]$Message = 'The requested user is invalid.') @{ status = 404; body = @{ error = @{ code = $Code; message = $Message } } } }
    if ($segments[0] -ne 'users') { return @{ status = 400; body = @{ error = @{ code = 'BadRequest'; message = "Unknown path $path" } } } }

    # ---- directory ------------------------------------------------------------------------------------------
    if ($segments.Count -eq 1) {
        $filter = [string]$q['$filter']
        $users = $script:Fake.Users
        if ($filter -match "proxyAddresses/any\(p:p eq 'smtp:(.+?)'\)") { $a = $Matches[1].ToLowerInvariant(); $users = @($users | Where-Object { @($_.proxyAddresses | ForEach-Object { $_.Substring(5).ToLowerInvariant() }) -contains $a }) }
        elseif ($filter -match "userPrincipalName eq '(.+?)'") { $a = $Matches[1].ToLowerInvariant(); $users = @($users | Where-Object { $_.userPrincipalName -eq $a }) }
        return @{ status = 200; body = @{ value = @($users | Select-Object id, displayName, mail, userPrincipalName) } }
    }
    if ($segments.Count -ge 3 -and $segments[2] -eq 'settings') {
        if ($Version -ne 'beta') { return @{ status = 400; body = @{ error = @{ code = 'BadRequest'; message = 'Resource not found for the segment settings/exchange.' } } } }
        $user = $script:Fake.Users | Where-Object { $_.id -eq $segments[1] -or $_.mail -eq $segments[1].ToLowerInvariant() } | Select-Object -First 1
        if (-not $user) { return (& $error404 'Request_ResourceNotFound' 'Resource not found.') }
        if ($user.OnPremises) { return (& $error404 'MailboxNotEnabledForRESTAPI' 'The mailbox is either inactive, soft-deleted, or is hosted on-premise.') }
        $body = [ordered]@{ primaryMailboxId = $user.PrimaryKey }
        if ($user.ArchiveKey) { $body.inPlaceArchiveMailboxId = $user.ArchiveKey }
        return @{ status = 200; body = $body }
    }

    # ---- mailboxes --------------------------------------------------------------------------------------------
    $store = $script:Fake.Mailboxes[$segments[1].ToLowerInvariant()]
    if (-not $store) {
        $user = $script:Fake.Users | Where-Object { $_.id -eq $segments[1] -or $_.mail -eq $segments[1].ToLowerInvariant() } | Select-Object -First 1
        if ($user -and $user.OnPremises) { return (& $error404 'MailboxNotEnabledForRESTAPI' 'The mailbox is either inactive, soft-deleted, or is hosted on-premise.') }
        return (& $error404)
    }
    # One message by its ID (the reading pane of the window): its body as text with Prefer outlook.body-content-type="text".
    if ($segments.Count -eq 4 -and $segments[2] -eq 'messages') {
        $msg = $store.Folders | ForEach-Object { $_.Messages } | Where-Object id -eq $segments[3] | Select-Object -First 1
        if (-not $msg) { return @{ status = 404; body = @{ error = @{ code = 'ErrorItemNotFound'; message = 'The specified object was not found in the store.' } } } }
        $content = [string]$msg.body.content
        $type = 'html'
        if ([string]$Headers['Prefer'] -match 'body-content-type="text"') { $content = ($content -replace '<br\s*/?>|</p>', "`r`n") -replace '<[^>]+>', ''; $type = 'text' }
        return @{ status = 200; body = [ordered]@{ id = $msg.id; body = @{ contentType = $type; content = $content } } }
    }
    if ($segments[2] -ne 'mailFolders') { return @{ status = 400; body = @{ error = @{ code = 'BadRequest'; message = "Unknown path $path" } } } }
    $page = {
        param([object[]]$Items, [int]$Size, [string]$Base)
        $skip = [int]$q['$skip']
        $values = @($Items | Select-Object -Skip $skip -First $Size)
        $body = [ordered]@{ value = $values }
        if ($skip + $Size -lt $Items.Count) {
            $rest = @($q.Keys | Where-Object { $_ -ne '$skip' } | ForEach-Object { "$([Uri]::EscapeDataString($_))=$([Uri]::EscapeDataString($q[$_]))" })
            $body['@odata.nextLink'] = "https://graph.microsoft.com/v1.0$($Base)?$(($rest + "`$skip=$($skip + $Size)") -join '&')"
        }
        elseif ($Base -match '/delta$') { $body['@odata.deltaLink'] = "https://graph.microsoft.com/v1.0$($Base)?`$deltatoken=end" }
        $body
    }
    $folderView = { param($f) [ordered]@{ id = $f.id; displayName = $f.displayName; parentFolderId = $f.parentFolderId; totalItemCount = $f.Messages.Count; childFolderCount = @($store.Folders | Where-Object parentFolderId -eq $f.id).Count } }
    if ($segments.Count -eq 4 -and $segments[3] -eq 'delta') {
        $size = 10
        if ($Headers['Prefer'] -match 'odata\.maxpagesize=(\d+)') { $size = [int]$Matches[1] }
        $all = @($store.Folders | Where-Object { -not $_.Recoverable } | ForEach-Object { & $folderView $_ })
        $cut = $false
        foreach ($pattern in @($script:Fake.CutOnce.Keys)) { if ($decoded -like $pattern -and $script:Fake.CutOnce[$pattern] -gt 0) { $script:Fake.CutOnce[$pattern]--; $cut = $true } }
        return @{ status = 200; body = (& $page $all $size $path); cut = $cut }
    }
    if ($segments.Count -eq 5 -and $segments[4] -eq 'childFolders') {
        $parentId = if ($segments[3] -eq 'recoverableitemsroot') { $store.RecoverableRoot } elseif ($segments[3] -eq 'msgfolderroot') { $store.Root } else { $segments[3] }
        $all = @($store.Folders | Where-Object parentFolderId -eq $parentId | ForEach-Object { & $folderView $_ })
        return @{ status = 200; body = (& $page $all ([Math]::Max(1, [int]$q['$top'])) $path) }
    }
    if ($segments.Count -eq 5 -and $segments[4] -eq 'messages') {
        $folder = $store.Folders | Where-Object id -eq $segments[3] | Select-Object -First 1
        if (-not $folder) { return @{ status = 404; body = @{ error = @{ code = 'ErrorItemNotFound'; message = 'The specified object was not found in the store.' } } } }
        $filter = [string]$q['$filter']
        $contains = $filter.IndexOf('contains(', [StringComparison]::OrdinalIgnoreCase)
        $date = $filter.IndexOf('receivedDateTime', [StringComparison]::OrdinalIgnoreCase)
        if ($contains -ge 0 -and ($date -lt 0 -or $contains -lt $date) -and $q.ContainsKey('$orderby')) { return @{ status = 400; body = @{ error = @{ code = 'InefficientFilter'; message = 'The restriction or sort order is too complex for this operation.' } } } }
        $items = @($folder.Messages | Where-Object { Test-FakeFilter $_ $filter } | Sort-Object { $_.receivedDateTime } -Descending:([string]$q['$orderby'] -notmatch ' asc$'))
        $select = @(([string]$q['$select']) -split ',' | Where-Object { $_ })
        $top = [Math]::Max(1, [int]$q['$top'])
        # A page cut short: once, larger than CutAbove, or holding a message Exchange cannot return.
        $cut = $false
        foreach ($pattern in @($script:Fake.CutOnce.Keys)) { if ($decoded -like $pattern -and $script:Fake.CutOnce[$pattern] -gt 0) { $script:Fake.CutOnce[$pattern]--; $cut = $true } }
        foreach ($pattern in @($script:Fake.CutAbove.Keys)) { if ($decoded -like $pattern -and $top -gt $script:Fake.CutAbove[$pattern]) { $cut = $true } }
        $sender = @($select | Where-Object { $_ -in 'from', 'sender', 'toRecipients', 'ccRecipients', 'bccRecipients' }).Count -gt 0
        foreach ($m in @($items | Select-Object -Skip ([int]$q['$skip']) -First $top)) {
            if ($m.PSObject.Properties['Cut'] -and ($m.Cut -eq 'Always' -or ($m.Cut -eq 'Recipients' -and $sender))) { $cut = $true }
        }
        # $select: the properties asked only (and id), as Graph answers.
        if ($select.Count) { $items = @($items | ForEach-Object { $o = $_; $view = [ordered]@{}; foreach ($p in $o.PSObject.Properties) { if ($p.Name -in $select -or $p.Name -in 'id', '@odata.type') { $view[$p.Name] = $p.Value } }; [pscustomobject]$view }) }
        return @{ status = 200; body = (& $page $items $top $path); cut = $cut }
    }
    return @{ status = 400; body = @{ error = @{ code = 'BadRequest'; message = "Unknown request $Method $path" } } }
}

function Invoke-FakeGraphHttp {
    <# What Start-MmrGraphSend would get back: @{ Status; RetryAfter; Content }. Handles $batch (v1.0 and beta). #>
    param([string]$Method, [string]$Url, [string]$Body, [hashtable]$Headers = @{})
    if ($Url -match '/\$batch$') {
        $version = if ($Url -match '/beta/') { 'beta' } else { 'v1.0' }
        $payload = $Body | ConvertFrom-Json -Depth 20
        $responses = foreach ($sub in @($payload.requests)) {
            $h = @{}
            if ($sub.PSObject.Properties['headers']) { foreach ($p in $sub.headers.PSObject.Properties) { $h[$p.Name] = [string]$p.Value } }
            $r = Invoke-FakeGraphRequest -Method $sub.method -Url $sub.url -Body $null -Headers $h -Version $version
            $item = [ordered]@{ id = $sub.id; status = $r.status }
            if ($r.ContainsKey('headers')) { $item.headers = $r.headers }
            if ($r.ContainsKey('body')) { $item.body = $r.body }
            $item
        }
        return @{ Status = 200; RetryAfter = 0; Content = (ConvertTo-FakeJson @{ responses = @($responses) }) }
    }
    $r = Invoke-FakeGraphRequest -Method $Method -Url $Url -Body $null -Headers $(if ($Headers) { $Headers } else { @{} })
    $retry = if ($r.ContainsKey('headers') -and $r.headers['Retry-After']) { [double]$r.headers['Retry-After'] } else { 0 }
    $content = if ($r.ContainsKey('body')) { ConvertTo-FakeJson $r.body } else { '' }
    # Cut short: the status and the start of the JSON, then nothing (as Exchange Online answers it).
    if ($r['cut']) { $content = $content.Substring(0, $content.IndexOf('"value":[') + 9) }
    return @{ Status = $r.status; RetryAfter = $retry; Content = $content }
}

function New-FakeToken {
    <# Unsigned app-only token with the roles given. #>
    param([string]$TenantId = '11111111-2222-3333-4444-555555555555', [string[]]$Roles = @('Mail.ReadBasic.All', 'User.Read.All'))
    $b64 = { param($o) [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes(($o | ConvertTo-Json -Compress))).TrimEnd('=').Replace('+', '-').Replace('/', '_') }
    '{0}.{1}.sig' -f (& $b64 @{ alg = 'none'; typ = 'JWT' }), (& $b64 @{ tid = $TenantId; roles = $Roles; app_displayname = 'Mailbox Message Report (test)'; appid = 'app' })
}
