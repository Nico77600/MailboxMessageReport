<#
.SYNOPSIS
    Mailbox Message Report - the mailboxes to read: user, primary mailbox, archive (dot-sourced by MailboxMessageReport.psm1).

.DESCRIPTION
    For each address of the request, in $batch calls:

      1. The user (User.Read.All or User.ReadBasic.All): by any SMTP address (proxyAddresses), then by UPN. Its ID is
         used to read the primary mailbox; without it, the address typed is used.
      2. Its mailboxes (User.Read.All): GET /beta/users/{id}/settings/exchange gives primaryMailboxId and, when the user
         has an archive, inPlaceArchiveMailboxId - both written MBX:<GUID>@<tenant ID>, the GUID being the ExchangeGuid
         or the ArchiveGuid of Exchange (checked against Get-EXOMailbox in the lab). No Exchange Online PowerShell.
         404 MailboxNotEnabledForRESTAPI: mailbox inactive, soft-deleted or on-premises - it cannot be read.
      3. The list of mailboxes may give the ArchiveGuid of a mailbox (CSV column ArchiveGuid, Get-EXOMailbox
         -Properties ArchiveGuid | Export-Csv): the archive is then MBX:<ArchiveGuid>@<tenant ID>, without User.Read.All.

    The archive is then read like a mailbox of its own: GET /v1.0/users/MBX:<ArchiveGuid>@<tenant ID>/mailFolders...
    (measured in the lab; Mail.ReadBasic.All is enough). An auto-expanding archive (auxiliary archives) is out of
    scope: a folder whose content moved to an auxiliary archive answers an error, shown in the report.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.0.0
#>

function Get-MmrUserPath {
    <# /users/<key>: a user ID or an MBX:<guid>@<tenant> as is, an address escaped. #>
    param([Parameter(Mandatory = $true)][string]$Key)
    if ($Key -match $script:GuidPattern -or $Key -match '^MBX:[0-9a-fA-F-]{36}@[0-9a-fA-F-]{36}$') { return "/users/$Key" }
    return "/users/$([Uri]::EscapeDataString($Key))"
}

function Get-MmrMailboxProblem {
    <# Why a mailbox cannot be read, in words, from a Graph answer (Status, ErrorCode, ErrorMessage). #>
    param([int]$Status, [string]$Code, [string]$Message)
    switch ($Code) {
        'MailboxNotEnabledForRESTAPI' { 'mailbox inactive, soft-deleted or on-premises' }
        'ErrorInvalidUser' { 'not a mailbox of this tenant' }
        'ResourceNotFound' { 'not a mailbox of this tenant' }
        'Request_ResourceNotFound' { 'not in the directory' }
        'ErrorNonExistentMailbox' { 'no mailbox' }
        'ErrorAccessDenied' { 'access denied (RBAC for Applications scope or application access policy)' }
        'Authorization_RequestDenied' { 'access denied by the directory (permission)' }
        default { "$Status $Code$(if ($Message) { ": $Message" })" }
    }
}

function Get-MmrMailboxState {
    <# NotFound | NotReachable | Denied | Error, from a Graph answer that refused a mailbox. #>
    param([int]$Status, [string]$Code)
    if ($Status -eq 403 -or $Code -in 'ErrorAccessDenied', 'Authorization_RequestDenied') { return 'Denied' }
    if ($Code -eq 'MailboxNotEnabledForRESTAPI') { return 'NotReachable' }
    if ($Status -eq 404) { return 'NotFound' }
    return 'Error'
}

function New-MmrMailbox {
    <# A mailbox of the run, with its counts (filled by the folders and the messages). #>
    param([Parameter(Mandatory = $true)][string]$Address, [string]$ArchiveGuid)
    [pscustomobject]@{
        Input = $Address; Address = $Address.ToLowerInvariant(); DisplayName = ''; UserId = ''; Key = $Address.ToLowerInvariant()
        PrimaryId = ''; ArchiveKey = ''; ArchiveGuid = [string]$ArchiveGuid; ArchiveSource = ''; Archive = 'Unknown'
        State = 'Ok'; Detail = ''; ArchiveState = ''; Notes = ''
        PrimaryFolders = 0; ArchiveFolders = 0; FoldersRead = 0; FoldersFailed = 0
        PrimaryMessages = [long]0; ArchiveMessages = [long]0; RecoverableMessages = [long]0; Messages = [long]0; Status = ''
        Csv = ''; Html = ''
    }
}

function Resolve-MmrMailboxes {
    <#
    .SYNOPSIS
        The mailboxes of the request: user, primary mailbox and archive of each address. Needs Connect-MmrGraph first.
    .PARAMETER Entries
        @{ Address; ArchiveGuid } per mailbox (New-MmrRequest).
    .OUTPUTS
        One object per mailbox (New-MmrMailbox): Address (primary SMTP when found), DisplayName, Key (the user ID, or the
        address), ArchiveKey (MBX:<ArchiveGuid>@<tenant>, '' without archive), Archive (Yes | No | Unknown), ArchiveSource
        (Graph | File), State (Ok | NotFound | NotReachable | Denied | Error) and Detail.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Entries)

    $g = $script:Graph
    $tenant = [string]$g.TenantGuid
    $list = [Collections.Generic.List[object]]::new()
    foreach ($e in $Entries) { $list.Add((New-MmrMailbox -Address $e.Address -ArchiveGuid $e.ArchiveGuid)) }
    if (-not $list.Count) { return , $list.ToArray() }
    $progress = { param($done, $total) Write-MmrProgress ($done / [Math]::Max(1, $total)) ('{0:N0}/{1:N0} mailboxes looked up' -f $done, $total) }

    # ---- 1. the user: any SMTP address, then the UPN ------------------------------------------------------
    if ($g.CanFindUsers) {
        $pending = $list.ToArray()
        foreach ($filter in "proxyAddresses/any(p:p eq 'smtp:{0}')", "userPrincipalName eq '{0}'") {
            if (-not $pending.Count) { break }
            $requests = for ($i = 0; $i -lt $pending.Count; $i++) {
                $f = [Uri]::EscapeDataString(($filter -f $pending[$i].Address.Replace("'", "''")))
                New-MmrGraphRequest -Id "u$i" -Url "/users?`$filter=$f&`$select=id,displayName,mail,userPrincipalName"
            }
            $res = Invoke-MmrGraphBatch -Requests @($requests) -OnProgress $progress
            $next = [Collections.Generic.List[object]]::new()
            for ($i = 0; $i -lt $pending.Count; $i++) {
                $m = $pending[$i]; $r = $res["u$i"]
                $user = if ($r.Status -eq 200) { $r.Values | Select-Object -First 1 } else { $null }
                if (-not $user) {
                    if ($r.Status -ne 200) { $m.Detail = "directory lookup: $(Get-MmrMailboxProblem $r.Status $r.ErrorCode $r.ErrorMessage)" }
                    $next.Add($m); continue
                }
                $m.UserId = [string]$user.id
                $m.Key = [string]$user.id
                $m.DisplayName = [string]$user.displayName
                if ($user.mail) { $m.Address = ([string]$user.mail).ToLowerInvariant() }
                $m.Detail = ''
            }
            $pending = $next.ToArray()
        }
        foreach ($m in $pending) { if (-not $m.Detail) { $m.Detail = 'not found in the directory: the address is read as typed' } }
    }

    # ---- 2. the mailboxes of each user: primary and archive (beta settings/exchange) --------------------------
    if ($g.CanReadUsers) {
        $found = @($list | Where-Object UserId)
        if ($found.Count) {
            $requests = for ($i = 0; $i -lt $found.Count; $i++) { New-MmrGraphRequest -Id "s$i" -Url "/users/$($found[$i].UserId)/settings/exchange" }
            $res = Invoke-MmrGraphBatch -Requests @($requests) -OnProgress $progress -Beta
            for ($i = 0; $i -lt $found.Count; $i++) {
                $m = $found[$i]; $r = $res["s$i"]
                if ($r.Status -eq 200) {
                    $m.PrimaryId = [string](Get-MmrProperty $r.Body 'primaryMailboxId')
                    $archive = [string](Get-MmrProperty $r.Body 'inPlaceArchiveMailboxId')
                    if ($archive) { $m.ArchiveKey = $archive; $m.Archive = 'Yes'; $m.ArchiveSource = 'Graph'; $m.ArchiveGuid = ($archive -replace '^MBX:([^@]+)@.*$', '$1').ToLowerInvariant() }
                    else { $m.Archive = 'No' }
                }
                elseif ($r.Status -eq 403) { $m.Notes = 'archive ID refused (settings/exchange 403): User.Read.All needed, or the ArchiveGuid in the list' }
                else {
                    $m.State = Get-MmrMailboxState $r.Status $r.ErrorCode
                    $m.Detail = Get-MmrMailboxProblem $r.Status $r.ErrorCode $r.ErrorMessage
                }
            }
        }
    }

    # ---- 3. the ArchiveGuid of the list, when Graph did not give the archive ---------------------------------
    foreach ($m in $list) {
        if ($m.Archive -ne 'Yes' -and $m.ArchiveGuid) {
            $m.ArchiveKey = "MBX:$($m.ArchiveGuid)@$tenant"
            $m.Archive = 'Yes'
            $m.ArchiveSource = 'File'
        }
        if ($m.State -eq 'Ok' -and -not $m.Detail) {
            $m.Detail = switch ($m.Archive) {
                'Yes' { if ($m.ArchiveSource -eq 'File') { 'archive from the ArchiveGuid of the list' } else { 'primary mailbox and archive' } }
                'No' { 'no archive' }
                default { if ($m.Notes) { $m.Notes } else { 'archive not known (User.Read.All, or the ArchiveGuid in the list)' } }
            }
        }
    }
    # An address typed twice under two aliases of one mailbox is read once.
    $unique = [Collections.Generic.List[object]]::new()
    $keys = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($m in $list) { if ($keys.Add($m.Key)) { $unique.Add($m) } }
    return , $unique.ToArray()
}
