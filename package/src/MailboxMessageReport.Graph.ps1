<#
.SYNOPSIS
    Mailbox Message Report - Microsoft Graph: token, requests, the $batch scheduler and the page reader
    (dot-sourced by MailboxMessageReport.psm1).

.DESCRIPTION
    Application permissions (app-only), as for every mailbox of the tenant (measured on a lab tenant, 2026-10-07):
      Mail.ReadBasic.All   required: the folders and the messages (subject, sender, recipients, dates, Internet
                           message ID) of the primary mailbox, of the archive and of Recoverable Items. Mail.Read or
                           Mail.ReadWrite count as well (they give the body too, which the tool never reads).
      User.Read.All        the user of each address (any alias), and the ID of its archive: GET /beta/users/{id}/
                           settings/exchange (inPlaceArchiveMailboxId = MBX:<ArchiveGuid>@<tenant ID>). User.ReadBasic.All
                           finds the user but the archive ID is refused (403). Without it, the archive is read only
                           when the list of mailboxes gives its ArchiveGuid.

    Token: certificate (client assertion built here, no module needed, recommended) or client secret
    (environment variable or typed, never written). It is renewed 5 minutes before it expires.

    Requests go through one transport (Start-MmrGraphSend / Complete-MmrGraphSend, replaced by a simulated
    tenant in the tests). Two schedulers use it:
      Invoke-MmrGraphBatch  many small requests (users, archive IDs) in $batch calls of 20, several calls at once;
      Invoke-MmrGraphPaged  lists read page after page (folders, messages), each page a request of its own (a page
                            of 1,000 messages is large), Graph.MaxConcurrency in flight.
    Both never send more than 4 requests at a time to the same mailbox (limit of Exchange Online; a primary mailbox
    and its archive are two mailboxes), retry 429 / 5xx after the Retry-After delay and renew the token on a 401.

.NOTES
    Author  : Nicolas Fabert
    Version : 2.1.0
#>

$script:GraphRoot = 'https://graph.microsoft.com/v1.0'
$script:GraphBeta = 'https://graph.microsoft.com/beta'
$script:LoginHost = 'https://login.microsoftonline.com'
$script:Http = $null
# Exchange Online processes at most 4 requests at a time for one mailbox and one application.
$script:MailboxConcurrency = 4
$script:BatchSize = 20

function ConvertTo-MmrBase64Url { param([byte[]]$Bytes) [Convert]::ToBase64String($Bytes).TrimEnd('=').Replace('+', '-').Replace('/', '_') }

function Get-MmrTokenClaims {
    <# Payload of a JWT (no signature check: only used to read tid, roles and the application name). #>
    param([Parameter(Mandatory = $true)][string]$Token)
    $parts = $Token.Split('.')
    if ($parts.Count -lt 2) { throw 'The access token is not a JWT.' }
    $p = $parts[1].Replace('-', '+').Replace('_', '/')
    while ($p.Length % 4) { $p += '=' }
    return [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($p)) | ConvertFrom-Json
}

function Get-MmrCertificate {
    <# Certificate with its private key, from Cert:\CurrentUser\My or Cert:\LocalMachine\My. #>
    param([Parameter(Mandatory = $true)][string]$Thumbprint)
    $thumb = $Thumbprint.Trim().ToUpperInvariant()
    foreach ($store in 'Cert:\CurrentUser\My', 'Cert:\LocalMachine\My') {
        $cert = Get-Item -LiteralPath (Join-Path $store $thumb) -ErrorAction SilentlyContinue
        if ($cert) {
            if (-not $cert.HasPrivateKey) { throw "Certificate $thumb found in $store without its private key: import the .pfx (not the .cer) for the account that runs the tool." }
            if ($cert.NotAfter -lt (Get-Date)) { throw "Certificate $thumb expired on $($cert.NotAfter.ToString('yyyy-MM-dd')). Upload a new certificate to the application and update Authentication.CertificateThumbprint." }
            return $cert
        }
    }
    throw "Certificate $thumb not found in Cert:\CurrentUser\My nor Cert:\LocalMachine\My (account $([Environment]::UserName)). Developer guide, chapter 5 'Application'."
}

function New-MmrClientAssertion {
    <# Client assertion (RFC 7523) signed with the certificate: RS256, header x5t = SHA-1 thumbprint, valid 10 minutes. #>
    param(
        [Parameter(Mandatory = $true)][Security.Cryptography.X509Certificates.X509Certificate2]$Certificate,
        [Parameter(Mandatory = $true)][string]$TenantId,
        [Parameter(Mandatory = $true)][string]$AppId
    )
    $now = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
    $header = [ordered]@{ alg = 'RS256'; typ = 'JWT'; x5t = (ConvertTo-MmrBase64Url $Certificate.GetCertHash()) } | ConvertTo-Json -Compress
    $claims = [ordered]@{ aud = "$($script:LoginHost)/$TenantId/oauth2/v2.0/token"; iss = $AppId; sub = $AppId; jti = [guid]::NewGuid().ToString(); nbf = $now - 60; iat = $now; exp = $now + 600 } | ConvertTo-Json -Compress
    $unsigned = (ConvertTo-MmrBase64Url ([Text.Encoding]::UTF8.GetBytes($header))) + '.' + (ConvertTo-MmrBase64Url ([Text.Encoding]::UTF8.GetBytes($claims)))
    $rsa = [Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($Certificate)
    if (-not $rsa) { throw "Certificate $($Certificate.Thumbprint): the private key is not an RSA key, or this account cannot use it." }
    try { $signature = $rsa.SignData([Text.Encoding]::ASCII.GetBytes($unsigned), [Security.Cryptography.HashAlgorithmName]::SHA256, [Security.Cryptography.RSASignaturePadding]::Pkcs1) }
    finally { $rsa.Dispose() }
    return "$unsigned.$(ConvertTo-MmrBase64Url $signature)"
}

function Get-MmrEntraErrorHint {
    <# A sentence for the most frequent Microsoft Entra sign-in errors. #>
    param([string]$Message)
    $hints = [ordered]@{
        'AADSTS700016'  = 'the application ID is not found in this tenant (Authentication.AppId, Tenant.TenantId).'
        'AADSTS90002'   = 'the tenant is not found (Tenant.TenantId).'
        'AADSTS700027'  = 'the certificate is not registered on the application, or not the right one (thumbprint).'
        'AADSTS7000215' = 'the client secret is not valid for this application.'
        'AADSTS7000222' = 'the client secret has expired: create a new one, or move to a certificate.'
        'AADSTS700024'  = 'the clock of this computer is not on time (the assertion is outside its validity).'
        'AADSTS53003'   = 'blocked by Conditional Access for workload identities.'
    }
    foreach ($code in $hints.Keys) { if ($Message -match $code) { return "$code - $($hints[$code])" } }
    return $null
}

function Get-MmrAppToken {
    <# App-only token (client credentials): certificate or secret, for Graph or another resource (Scope). Returns @{ Token; ExpiresUtc }. #>
    param([Parameter(Mandatory = $true)][hashtable]$Settings, $Certificate, [Security.SecureString]$Secret, [string]$Scope = 'https://graph.microsoft.com/.default')
    $body = @{ client_id = $Settings.AppId; scope = $Scope; grant_type = 'client_credentials' }
    if ($Certificate) {
        $body['client_assertion_type'] = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'
        $body['client_assertion'] = New-MmrClientAssertion -Certificate $Certificate -TenantId $Settings.TenantId -AppId $Settings.AppId
    }
    else { $body['client_secret'] = [Net.NetworkCredential]::new('', $Secret).Password }
    try { $r = Invoke-RestMethod -Method Post -Uri "$($script:LoginHost)/$($Settings.TenantId)/oauth2/v2.0/token" -Body $body -ErrorAction Stop }
    catch {
        $text = try { ($_.ErrorDetails.Message | ConvertFrom-Json).error_description } catch { $null }
        if (-not $text) { $text = $_.Exception.Message }
        $hint = Get-MmrEntraErrorHint $text
        throw ('Microsoft Entra sign-in of the application failed{0}: {1}' -f $(if ($hint) { " ($hint)" } else { '' }), ($text -split "`r?`n")[0])
    }
    finally { $body.Clear() }
    return @{ Token = $r.access_token; ExpiresUtc = [datetime]::UtcNow.AddSeconds([int]$r.expires_in) }
}

function Connect-MmrGraph {
    <#
    .SYNOPSIS
        Obtains the first token, checks the tenant and the permissions, and keeps the connection for the run.
    .PARAMETER Secret
        ClientSecret mode: the secret (window). Otherwise the environment variable, else a prompt.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][hashtable]$Settings, [Security.SecureString]$Secret)

    $check = Test-MmrConfiguration -Configuration $Settings -ForConnection
    if (-not $check.IsValid) { throw ("Connection settings:`n - " + ($check.Problems -join "`n - ")) }
    $connection = @{ TenantId = $Settings.TenantId; AppId = $Settings.AppId; Mode = $Settings.AuthMode; Settings = $Settings; Certificate = $null; Secret = $null; Token = $null; ExpiresUtc = [datetime]::MinValue }
    if ($Settings.AuthMode -eq 'Certificate') {
        $connection.Certificate = Get-MmrCertificate $Settings.CertificateThumbprint
        if ($connection.Certificate.NotAfter -lt (Get-Date).AddDays(30)) { Write-MmrItem Warn ('Certificate {0} expires on {1}: renew it.' -f $connection.Certificate.Thumbprint, $connection.Certificate.NotAfter.ToString('yyyy-MM-dd')) }
    }
    else {
        if (-not $Secret) {
            $value = [Environment]::GetEnvironmentVariable($Settings.ClientSecretVariable)
            if ($value) { $Secret = ConvertTo-SecureString $value -AsPlainText -Force; $value = $null }
            elseif (-not $script:Ui -and [Environment]::UserInteractive -and -not [Console]::IsInputRedirected) { $Secret = Read-Host -AsSecureString "      Client secret of application $($Settings.AppId)" }
            else { throw "ClientSecret mode: the environment variable $($Settings.ClientSecretVariable) is empty and no secret was typed." }
        }
        if (-not $Secret -or $Secret.Length -eq 0) { throw 'No client secret given.' }
        $connection.Secret = $Secret
    }
    $script:Graph = $connection
    Update-MmrToken -Force

    $claims = Get-MmrTokenClaims $connection.Token
    if ([string]$Settings.TenantId -match $script:GuidPattern -and $claims.tid -ne $Settings.TenantId) {
        $script:Graph = $null
        throw "Connected to tenant $($claims.tid), but Tenant.TenantId is $($Settings.TenantId). Nothing was read."
    }
    $roles = @($claims.PSObject.Properties['roles'] | ForEach-Object { $_.Value })
    $connection.Roles = $roles
    $connection.TenantGuid = [string]$claims.tid
    $connection.AppName = [string](Get-MmrProperty $claims 'app_displayname')
    # Lab, 2026-10-07: Mail.ReadBasic.All reads the folders and every property of the report (subject, from, to, cc,
    # bcc, dates, Internet message ID) in the primary mailbox, the archive and Recoverable Items; Mail.Read and
    # Mail.ReadWrite read them too (and the body, never read here).
    $connection.CanReadMail = @($roles | Where-Object { $_ -in 'Mail.ReadBasic.All', 'Mail.ReadBasic', 'Mail.Read', 'Mail.ReadWrite' }).Count -gt 0
    # The archive ID (settings/exchange) needs User.Read.All: refused (403) with User.ReadBasic.All.
    $connection.CanReadUsers = @($roles | Where-Object { $_ -in 'User.Read.All', 'User.ReadWrite.All', 'Directory.Read.All', 'Directory.ReadWrite.All' }).Count -gt 0
    $connection.CanFindUsers = $connection.CanReadUsers -or $roles -contains 'User.ReadBasic.All'
    $grant = "Entra admin center > App registrations > $(if ($connection.AppName) { $connection.AppName } else { $Settings.AppId }) > API permissions > Microsoft Graph > Application permissions, then 'Grant admin consent'"
    if (-not $connection.CanReadMail) {
        $script:Graph = $null
        throw "The application has no application permission Mail.ReadBasic.All (or Mail.Read) with admin consent (roles in the token: $(if ($roles.Count) { $roles -join ', ' } else { 'none' })). $grant."
    }
    $connection.Grant = $grant
    return [pscustomobject]$connection
}

function Update-MmrToken {
    <# Renews the token when it expires within 5 minutes (or at once with -Force, after a 401). #>
    param([switch]$Force)
    $g = $script:Graph
    if (-not $g) { throw 'Not connected to Microsoft Graph (Connect-MmrGraph).' }
    if (-not $Force -and $g.Token -and $g.ExpiresUtc -gt [datetime]::UtcNow.AddMinutes(5)) { return }
    if ($g.ContainsKey('Renew') -and $g.Renew) { $t = & $g.Renew }
    else { $t = Get-MmrAppToken -Settings $g.Settings -Certificate $g.Certificate -Secret $g.Secret }
    $g.Token = $t.Token
    $g.ExpiresUtc = $t.ExpiresUtc
    Write-MmrLog 'INFO' ('Access token obtained, valid until {0:HH:mm:ss} UTC.' -f $t.ExpiresUtc)
}

function Get-MmrHttpClient {
    if (-not $script:Http) {
        $handler = [Net.Http.SocketsHttpHandler]::new()
        $handler.MaxConnectionsPerServer = 32
        $handler.AutomaticDecompression = [Net.DecompressionMethods]::GZip -bor [Net.DecompressionMethods]::Deflate
        $script:Http = [Net.Http.HttpClient]::new($handler)
        $script:Http.Timeout = [TimeSpan]::FromSeconds($(if ($script:Graph) { [int]$script:Graph.Settings.TimeoutSeconds } else { 120 }))
        $script:Http.DefaultRequestHeaders.UserAgent.ParseAdd("MailboxMessageReport/$($script:ToolVersion)")
    }
    return $script:Http
}

function Start-MmrGraphSend {
    <#
        Sends one HTTP request to Graph without waiting. Returns a handle for Complete-MmrGraphSend.
        The only function that touches the network for Graph: the tests replace it with a simulated tenant.
    #>
    param([Parameter(Mandatory = $true)][string]$Method, [Parameter(Mandatory = $true)][string]$Url, [string]$Body, [hashtable]$Headers)
    $request = [Net.Http.HttpRequestMessage]::new([Net.Http.HttpMethod]::new($Method), $Url)
    $request.Headers.Authorization = [Net.Http.Headers.AuthenticationHeaderValue]::new('Bearer', $script:Graph.Token)
    [void]$request.Headers.TryAddWithoutValidation('client-request-id', [guid]::NewGuid().ToString())
    if ($Headers) { foreach ($h in $Headers.Keys) { [void]$request.Headers.TryAddWithoutValidation([string]$h, [string]$Headers[$h]) } }
    if ($Body) { $request.Content = [Net.Http.StringContent]::new($Body, [Text.Encoding]::UTF8, 'application/json') }
    [pscustomobject]@{ Task = (Get-MmrHttpClient).SendAsync($request); Request = $request; Response = $null }
}

function Test-MmrGraphSendDone { param($Handle) return ($null -ne $Handle.Response) -or $Handle.Task.IsCompleted }

function Complete-MmrGraphSend {
    <# Result of a sent request: @{ Status; RetryAfter; Content (Body: bytes, its text on demand) }. Status 0 = no answer (network, timeout). #>
    param([Parameter(Mandatory = $true)]$Handle)
    if ($null -ne $Handle.Response) { return $Handle.Response }
    $response = $null
    try {
        $response = $Handle.Task.GetAwaiter().GetResult()
        # The body as bytes (compiled): a page of messages is parsed from them, never passed as a PowerShell string.
        $content = [MailboxMessageReportNative.Body]::Read($response.Content)
        $retry = if ($response.Headers.RetryAfter -and $response.Headers.RetryAfter.Delta) { $response.Headers.RetryAfter.Delta.Value.TotalSeconds } else { 0 }
        return @{ Status = [int]$response.StatusCode; RetryAfter = $retry; Content = $content }
    }
    catch {
        $e = $_.Exception
        while ($e.InnerException) { $e = $e.InnerException }
        return @{ Status = 0; RetryAfter = 0; Content = ''; Error = $e.Message }
    }
    finally {
        if ($response) { $response.Dispose() }
        if ($Handle.Request) { $Handle.Request.Dispose() }
    }
}

function Wait-MmrUi {
    <# Waits a little while keeping the window responsive; throws when the user stopped the run. #>
    param([int]$Milliseconds = 50, [switch]$NoCancel)
    if ($script:Ui -and $script:Ui.Pump) {
        $until = [datetime]::UtcNow.AddMilliseconds($Milliseconds)
        do { & $script:Ui.Pump; Start-Sleep -Milliseconds 15 } while ([datetime]::UtcNow -lt $until)
    }
    else { Start-Sleep -Milliseconds $Milliseconds }
    if (-not $NoCancel) { Assert-MmrNotCancelled }
}

function Assert-MmrNotCancelled {
    if ($script:Ui -and $script:Ui.Cancel) { throw [OperationCanceledException]::new('Stopped by the user.') }
}

function Get-MmrGraphError {
    <# Code and message of a Graph error body. #>
    param($Body)
    $code = ''; $message = ''
    if ($Body -and $Body.PSObject.Properties['error']) {
        $code = [string]$Body.error.code
        $message = [string]$Body.error.message
    }
    return @{ Code = $code; Message = $message }
}

function ConvertFrom-MmrJson {
    param([AllowEmptyString()][AllowNull()][string]$Content)
    if ([string]::IsNullOrWhiteSpace($Content)) { return $null }
    try { return $Content | ConvertFrom-Json -Depth 64 } catch { return [pscustomobject]@{ error = [pscustomobject]@{ code = 'InvalidJson'; message = ($Content.Substring(0, [Math]::Min(300, $Content.Length))) } } }
}

function Get-MmrMailboxKey {
    <#
        The mailbox a request reads (/users/<address, user ID or MBX:guid@tenant>/...), to keep at most 4 requests at a
        time per mailbox. A primary mailbox and its archive are two mailboxes for Exchange Online.
    #>
    param([string]$Url)
    $m = [regex]::Match($Url, '^(?:https://graph\.microsoft\.com/(?:v1\.0|beta))?/?users/([^/?]+)', 'IgnoreCase')
    if ($m.Success) { return [Uri]::UnescapeDataString($m.Groups[1].Value).ToLowerInvariant() }
    return '(directory)'
}

function Invoke-MmrGraph {
    <#
        One request (not batched), with the same retries as the batches. Path relative to /v1.0 or an
        absolute URL (nextLink). Returns @{ Status; Body; ErrorCode; ErrorMessage }. Headers: Prefer...
    #>
    param([string]$Method = 'GET', [Parameter(Mandatory = $true)][string]$Path, $Body, [hashtable]$Headers)

    $url = if ($Path -match '^https://') { $Path } else { "$($script:GraphRoot)/$($Path.TrimStart('/'))" }
    $json = if ($null -eq $Body) { $null } elseif ($Body -is [string]) { $Body } else { $Body | ConvertTo-Json -Depth 20 -Compress }
    $max = [int]$script:Graph.Settings.MaxRetries
    for ($attempt = 0; ; $attempt++) {
        Update-MmrToken
        $handle = Start-MmrGraphSend -Method $Method -Url $url -Body $json -Headers $Headers
        while (-not (Test-MmrGraphSendDone $handle)) { Wait-MmrUi 30 -NoCancel }
        $r = Complete-MmrGraphSend $handle
        $parsed = ConvertFrom-MmrJson $r.Content
        if ($r.Status -eq 401 -and $attempt -lt 1) { Update-MmrToken -Force; continue }
        if ($r.Status -in 0, 429, 500, 502, 503, 504 -and $attempt -lt $max) {
            $delay = if ($r.RetryAfter -gt 0) { $r.RetryAfter } else { [Math]::Min(60, [Math]::Pow(2, $attempt + 1)) }
            Write-MmrLog 'WARN' ("Graph $Method $Path -> $($r.Status)$(if ($r['Error']) { " ($($r['Error']))" }), retry in $delay s")
            Wait-MmrUi ([int]($delay * 1000))
            continue
        }
        $err = Get-MmrGraphError $parsed
        if ($r.Status -eq 0) { $err = @{ Code = 'NoResponse'; Message = [string]$r['Error'] } }
        return [pscustomobject]@{ Status = $r.Status; Body = $parsed; ErrorCode = $err.Code; ErrorMessage = $err.Message }
    }
}

function Get-MmrGraphAll {
    <# Every item of a list (follows @odata.nextLink). Throws on an error, with the Graph message. #>
    param([Parameter(Mandatory = $true)][string]$Path)
    $items = [Collections.Generic.List[object]]::new()
    $next = $Path
    while ($next) {
        $r = Invoke-MmrGraph -Path $next
        if ($r.Status -ne 200) { throw "Graph GET $Path -> $($r.Status) $($r.ErrorCode): $($r.ErrorMessage)" }
        foreach ($v in @($r.Body.value)) { if ($null -ne $v) { $items.Add($v) } }
        $next = if ($r.Body.PSObject.Properties['@odata.nextLink']) { [string]$r.Body.'@odata.nextLink' } else { $null }
    }
    return , $items.ToArray()
}

function New-MmrGraphRequest {
    <# One request for Invoke-MmrGraphBatch. Url relative to /v1.0 (/users/...). Headers: Prefer (immutable IDs)... #>
    param([Parameter(Mandatory = $true)][string]$Id, [Parameter(Mandatory = $true)][string]$Url, [string]$Method = 'GET', $Body, [hashtable]$Headers)
    [pscustomobject]@{ Id = $Id; Method = $Method; Url = '/' + $Url.TrimStart('/'); Body = $Body; Headers = $Headers; Mailbox = (Get-MmrMailboxKey $Url.TrimStart('/')) }
}

function Invoke-MmrGraphBatch {
    <#
    .SYNOPSIS
        Sends many requests through $batch and returns their results by Id.
    .PARAMETER Requests
        Objects of New-MmrGraphRequest (Id unique).
    .PARAMETER FollowPages
        A list answer with @odata.nextLink is followed: Values holds every item of every page.
    .PARAMETER OnProgress
        Called after each $batch call with (requests done, requests in total).
    .PARAMETER Beta
        The requests are sent to the beta endpoint of Graph (settings/exchange: the ID of the archive).
    .OUTPUTS
        Hashtable Id -> [pscustomobject]@{ Id; Status; Body; Values; ErrorCode; ErrorMessage; Done; DoneUtc }.
        Stopped by the user: OperationCanceledException, with the results so far in Exception.Data['Results'].
    #>
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Requests,
        [switch]$FollowPages,
        [scriptblock]$OnProgress,
        [switch]$Beta
    )

    $results = @{}
    if (-not $Requests.Count) { return $results }
    $settings = $script:Graph.Settings
    $root = if ($Beta) { $script:GraphBeta } else { $script:GraphRoot }
    $maxRetries = [int]$settings.MaxRetries
    $concurrency = [Math]::Max(1, [int]$settings.MaxConcurrency)
    $queue = [Collections.Generic.List[object]]::new()
    foreach ($r in $Requests) {
        $queue.Add(@{ Request = $r; Attempt = 0; NotBefore = [datetime]::MinValue })
        $results[$r.Id] = [pscustomobject]@{ Id = $r.Id; Status = 0; Body = $null; Values = [Collections.Generic.List[object]]::new(); ErrorCode = ''; ErrorMessage = ''; Done = $false; DoneUtc = [datetime]::MinValue }
    }
    $inflight = [Collections.Generic.List[object]]::new()
    $busy = @{}
    $total = $Requests.Count; $done = 0; $calls = 0; $retries = 0
    $started = [datetime]::UtcNow
    $requeue = {
        param($item, [double]$Seconds)
        $item.Attempt++
        $item.NotBefore = [datetime]::UtcNow.AddSeconds($Seconds)
        $queue.Add($item)
    }
    $finish = {
        param($item, [int]$Status, $Body)
        $res = $results[$item.Request.Id]
        $res.Status = $Status
        $err = Get-MmrGraphError $Body
        $res.ErrorCode = $err.Code; $res.ErrorMessage = $err.Message
        if ($Status -eq 200 -and $Body -and $Body.PSObject.Properties['value']) {
            foreach ($v in @($Body.value)) { if ($null -ne $v) { $res.Values.Add($v) } }
            $next = if ($Body.PSObject.Properties['@odata.nextLink']) { [string]$Body.'@odata.nextLink' } else { '' }
            if ($FollowPages -and $next) {
                $nextUrl = $next -replace '^https://graph\.microsoft\.com/(v1\.0|beta)', ''
                $queue.Add(@{ Request = [pscustomobject]@{ Id = $item.Request.Id; Method = 'GET'; Url = $nextUrl; Body = $null; Headers = $item.Request.Headers; Mailbox = $item.Request.Mailbox }; Attempt = 0; NotBefore = [datetime]::MinValue })
                return
            }
        }
        $res.Body = $Body
        $res.Done = $true
        $res.DoneUtc = [datetime]::UtcNow
        $script:MmrBatchDone++
    }

    $script:MmrBatchDone = 0
    $cancelled = $false
    while ($queue.Count -or $inflight.Count) {
        if (-not $cancelled -and $script:Ui -and $script:Ui.Cancel) { $cancelled = $true }
        if ($cancelled -and -not $inflight.Count) {
            $stop = [OperationCanceledException]::new('Stopped by the user.')
            $stop.Data['Results'] = $results
            throw $stop
        }

        # ---- schedule new $batch calls --------------------------------------------------------
        while (-not $cancelled -and $inflight.Count -lt $concurrency -and $queue.Count) {
            $now = [datetime]::UtcNow
            $picked = [Collections.Generic.List[object]]::new()
            $inBatch = @{}
            foreach ($item in $queue) {
                if ($picked.Count -ge $script:BatchSize) { break }
                if ($item.NotBefore -gt $now) { continue }
                $key = $item.Request.Mailbox
                $used = [int]$busy[$key] + [int]$inBatch[$key]
                # The limit of 4 applies to a mailbox; directory requests (users, groups, places) are not limited.
                if ($used -ge $script:MailboxConcurrency -and -not $key.StartsWith('(')) { continue }
                $inBatch[$key] = [int]$inBatch[$key] + 1
                $picked.Add($item)
            }
            if (-not $picked.Count) { break }
            foreach ($item in $picked) { [void]$queue.Remove($item); $busy[$item.Request.Mailbox] = [int]$busy[$item.Request.Mailbox] + 1 }
            $subs = for ($i = 0; $i -lt $picked.Count; $i++) {
                $req = $picked[$i].Request
                $sub = [ordered]@{ id = [string]$i; method = $req.Method; url = $req.Url }
                $headers = @{}
                if ($req.PSObject.Properties['Headers'] -and $req.Headers) { foreach ($h in $req.Headers.Keys) { $headers[$h] = $req.Headers[$h] } }
                if ($null -ne $req.Body) { $headers['Content-Type'] = 'application/json'; $sub.body = $req.Body }
                if ($headers.Count) { $sub.headers = $headers }
                $sub
            }
            Update-MmrToken
            $payload = @{ requests = @($subs) } | ConvertTo-Json -Depth 20 -Compress
            $inflight.Add(@{ Handle = (Start-MmrGraphSend -Method 'POST' -Url "$root/`$batch" -Body $payload); Items = $picked })
            $calls++
        }

        if (-not $inflight.Count) {
            # Only requests waiting for their Retry-After.
            $wait = ($queue | ForEach-Object { $_.NotBefore } | Measure-Object -Minimum).Minimum
            $ms = [int][Math]::Max(50, [Math]::Min(5000, ($wait - [datetime]::UtcNow).TotalMilliseconds))
            Wait-MmrUi $ms -NoCancel
            continue
        }

        # ---- wait for a $batch call to finish ------------------------------------------------
        $finished = @($inflight | Where-Object { Test-MmrGraphSendDone $_.Handle })
        if (-not $finished.Count) { Wait-MmrUi 25 -NoCancel; continue }
        foreach ($call in $finished) {
            [void]$inflight.Remove($call)
            foreach ($item in $call.Items) { $busy[$item.Request.Mailbox] = [int]$busy[$item.Request.Mailbox] - 1 }
            $response = Complete-MmrGraphSend $call.Handle
            if ($response.Status -eq 200) {
                $parsed = ConvertFrom-MmrJson $response.Content
                $byId = @{}
                foreach ($sub in @(Get-MmrProperty $parsed 'responses')) { if ($sub) { $byId[[string]$sub.id] = $sub } }
                for ($i = 0; $i -lt $call.Items.Count; $i++) {
                    $item = $call.Items[$i]
                    $sub = $byId[[string]$i]
                    if (-not $sub) { & $requeue $item 2; $retries++; continue }
                    $status = [int]$sub.status
                    $body = if ($sub.PSObject.Properties['body']) { $sub.body } else { $null }
                    if ($status -in 429, 500, 502, 503, 504 -and $item.Attempt -lt $maxRetries) {
                        $after = 0.0
                        if ($sub.PSObject.Properties['headers'] -and $sub.headers -and $sub.headers.PSObject.Properties['Retry-After']) { [void][double]::TryParse([string]$sub.headers.'Retry-After', [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$after) }
                        if ($after -le 0) { $after = [Math]::Min(60, [Math]::Pow(2, $item.Attempt + 1)) }
                        & $requeue $item $after; $retries++
                        continue
                    }
                    if ($status -eq 401 -and $item.Attempt -lt 1) { Update-MmrToken -Force; & $requeue $item 0; $retries++; continue }
                    & $finish $item $status $body
                }
            }
            elseif ($response.Status -eq 401 -and @($call.Items | Where-Object { $_.Attempt -lt 1 }).Count) {
                Update-MmrToken -Force
                foreach ($item in $call.Items) { & $requeue $item 0 }
                $retries += $call.Items.Count
            }
            elseif ($response.Status -in 0, 429, 500, 502, 503, 504 -and @($call.Items | Where-Object { $_.Attempt -lt $maxRetries }).Count) {
                $after = if ($response.RetryAfter -gt 0) { $response.RetryAfter } else { [Math]::Min(60, [Math]::Pow(2, $call.Items[0].Attempt + 1)) }
                Write-MmrLog 'WARN' ("Graph `$batch -> $($response.Status)$(if ($response['Error']) { " ($($response['Error']))" }), $($call.Items.Count) request(s) retried in $after s")
                foreach ($item in $call.Items) { & $requeue $item $after }
                $retries += $call.Items.Count
            }
            else {
                $parsed = ConvertFrom-MmrJson $response.Content
                if ($response.Status -eq 0) { $parsed = [pscustomobject]@{ error = [pscustomobject]@{ code = 'NoResponse'; message = [string]$response['Error'] } } }
                foreach ($item in $call.Items) { & $finish $item $response.Status $parsed }
            }
        }
        if ($OnProgress) { & $OnProgress $script:MmrBatchDone $total }
    }
    Write-MmrLog 'INFO' ("Graph: {0} request(s) in {1} `$batch call(s), {2} retried, {3}" -f $total, $calls, $retries, (Format-MmrDuration ([datetime]::UtcNow - $started).TotalSeconds))
    foreach ($res in $results.Values) { if ($res.Status -eq 200 -and $null -eq $res.Body) { $res.Body = [pscustomobject]@{ value = $res.Values.ToArray() } } }
    return $results
}

function New-MmrPagedJob {
    <# One list to read page after page (Invoke-MmrGraphPaged). Url relative to /v1.0, or absolute; Tag: anything for the caller. #>
    param([Parameter(Mandatory = $true)][string]$Id, [Parameter(Mandatory = $true)][string]$Url, [hashtable]$Headers, $Tag)
    $u = if ($Url -match '^https://') { $Url } else { "$($script:GraphRoot)/$($Url.TrimStart('/'))" }
    # PageSize, Select: as asked, restored after a page cut short (Step-MmrPagedCut). Skipped, Partial: messages left
    # out, or read without their sender and recipients, because Graph could not return them.
    [pscustomobject]@{
        Id = $Id; Url = $u; Headers = $Headers; Tag = $Tag; Mailbox = (Get-MmrMailboxKey ($u -replace '^https://graph\.microsoft\.com/(v1\.0|beta)', '')); Attempt = 0; NotBefore = [datetime]::MinValue; Pages = 0
        PageSize = [int](Get-MmrUrlParameter $u 'top'); Select = [string](Get-MmrUrlParameter $u 'select'); Cuts = 0; CutTotal = 0; Reduced = $false; Skipped = 0; Partial = 0
        Started = $null; Seconds = [double]0
    }
}

function Get-MmrUrlParameter {
    <# A parameter of the query of a URL ($top, %24top...), decoded; $null when absent. Name without the $. #>
    param([string]$Url, [string]$Name)
    $m = [regex]::Match($Url, "[?&](?:\`$|%24)$Name=([^&]*)", 'IgnoreCase')
    if ($m.Success) { return [Uri]::UnescapeDataString($m.Groups[1].Value) }
    return $null
}

function Set-MmrUrlParameter {
    <# The URL with a parameter of its query set ($top, $skip, $select...), added when absent. Name without the $. #>
    param([string]$Url, [string]$Name, [string]$Value)
    $text = "`$$Name=$([Uri]::EscapeDataString($Value))"
    $re = [regex]::new("([?&])(?:\`$|%24)$Name=[^&]*", 'IgnoreCase')
    if ($re.IsMatch($Url)) { return $re.Replace($Url, '${1}' + $text.Replace('$', '$$'), 1) }
    return "$Url$(if ($Url.Contains('?')) { '&' } else { '?' })$text"
}

function Get-MmrCutAnswer {
    <# The reason when an exception says that a page was not complete JSON (the answer of Graph cut short), else $null. #>
    param([Exception]$Exception)
    for ($e = $Exception; $e; $e = $e.InnerException) {
        if ($e -is [System.Text.Json.JsonException] -or $e -is [IO.InvalidDataException]) { return $e.Message }
    }
    return $null
}

function Step-MmrPagedCut {
    <#
    .SYNOPSIS
        A page answered 200 whose JSON is cut short: what to ask next. Returns @{ Retry; Delay; Message }.
    .DESCRIPTION
        Exchange Online sometimes ends the answer of a page of messages early (the status line and the start of the
        JSON are already sent): a page that took it too long (the recipients of meeting messages are read one by one),
        or a message it cannot return. The page is asked again once; then with half as many messages, down to one;
        then that one message without its sender and recipients; then it is left out ($skip) and the list goes on.
        Every step is logged; Skipped and Partial count the messages lost or incomplete. Other lists (folders) are
        asked again up to Graph.MaxRetries times.
    #>
    param([Parameter(Mandatory = $true)]$Job, [long]$Length, [string]$Reason, [int]$MaxRetries)
    $Job.Cuts++
    $Job.CutTotal++
    # The reason (the error of the JSON reader) once per list: the next lines of the same list are shorter.
    $what = "Graph answered a page cut short ($('{0:N0}' -f $Length) bytes$(if ($Job.CutTotal -eq 1) { ": $Reason" }))"
    $log = { param([string]$Text) Write-MmrLog 'WARN' "$what - $Text - $($Job.Url)" }
    if ($Job.CutTotal -gt 500) { return @{ Retry = $false; Message = "$what, $($Job.CutTotal) times" } }
    if (-not ($Job.PageSize -gt 0 -and $Job.Url -match '/messages\?')) {
        if ($Job.Cuts -le $MaxRetries) { & $log "asked again ($($Job.Cuts))"; return @{ Retry = $true; Delay = [Math]::Min(30, [Math]::Pow(2, $Job.Cuts - 1)) } }
        return @{ Retry = $false; Message = "$what, $($Job.Cuts) times" }
    }
    # The first time on this list: perhaps once only.
    if ($Job.CutTotal -eq 1) { & $log 'asked again'; return @{ Retry = $true; Delay = 1 } }
    $top = [int](Get-MmrUrlParameter $Job.Url 'top')
    if ($top -gt 1) {
        $half = [int][Math]::Floor($top / 2)
        $Job.Url = Set-MmrUrlParameter $Job.Url 'top' $half
        & $log "asked again with $half message(s) a page"
        return @{ Retry = $true; Delay = 0 }
    }
    $select = @(([string](Get-MmrUrlParameter $Job.Url 'select')) -split ',')
    if (-not $Job.Reduced -and @($select | Where-Object { $_ -in 'from', 'sender', 'toRecipients', 'ccRecipients', 'bccRecipients' }).Count) {
        $Job.Url = Set-MmrUrlParameter $Job.Url 'select' (@($select | Where-Object { $_ -notin 'from', 'sender', 'toRecipients', 'ccRecipients', 'bccRecipients' }) -join ',')
        $Job.Reduced = $true
        & $log 'one message: asked again without its sender and recipients'
        return @{ Retry = $true; Delay = 0 }
    }
    if ($null -ne (Get-MmrUrlParameter $Job.Url 'skiptoken')) { return @{ Retry = $false; Message = "$what, even for one message" } }
    $skip = [int](Get-MmrUrlParameter $Job.Url 'skip')
    $Job.Url = Set-MmrUrlParameter (Set-MmrUrlParameter $Job.Url 'skip' ($skip + 1)) 'top' $Job.PageSize
    if ($Job.Select) { $Job.Url = Set-MmrUrlParameter $Job.Url 'select' $Job.Select }
    $Job.Reduced = $false
    $Job.Cuts = 0
    $Job.Skipped++
    & $log "message $($skip + 1) of the list left out: Graph cannot return it"
    return @{ Retry = $true; Delay = 0 }
}

function Get-MmrPagedNext {
    <# After a page read on a list that had pages cut short: the next page with the $select asked, and a larger page again. #>
    param([Parameter(Mandatory = $true)]$Job, [string]$Next)
    if ($Job.Reduced) {
        $Job.Partial++
        Write-MmrLog 'WARN' "A message read without its sender and recipients (Graph cannot return them) - $($Job.Url)"
        $Job.Reduced = $false
        if ($Next -and $Job.Select) { $Next = Set-MmrUrlParameter $Next 'select' $Job.Select }
    }
    $Job.Cuts = 0
    if ($Next -and $Job.PageSize -gt 0) {
        $top = [int](Get-MmrUrlParameter $Next 'top')
        if ($top -gt 0 -and $top -lt $Job.PageSize) { $Next = Set-MmrUrlParameter $Next 'top' ([Math]::Min($Job.PageSize, $top * 2)) }
    }
    return $Next
}

function Invoke-MmrGraphPaged {
    <#
    .SYNOPSIS
        Reads many lists (the folders of a mailbox, the messages of a folder) page after page, several at once.
    .DESCRIPTION
        Each page is a request of its own (not $batch: a page of 1,000 messages is about 1.5 MB). At most
        Graph.MaxConcurrency requests in flight, and at most 4 at a time for one mailbox: the lists of each mailbox
        wait in a queue of their own and the mailboxes are served in turn, so that hundreds of mailboxes and thousands
        of folders are scheduled at the same cost as a few. The next page of a list goes before the other lists of its
        mailbox (a folder is finished before the next one starts). 429 and 5xx (and no answer) are retried after the
        Retry-After delay, up to Graph.MaxRetries times per page; a 401 renews the token once.
    .PARAMETER Jobs
        Objects of New-MmrPagedJob.
    .PARAMETER OnPage
        { param($Job, $Content) } for each page answered 200 (the body as received: a Body, its JSON text with [string]);
        returns the nextLink to follow, or '' / $null after the last page.
    .PARAMETER OnDone
        { param($Job, [int]$Status, [string]$Code, [string]$Message) } once per job: 200 when every page was read,
        else the error of the page that failed (the pages before were given to OnPage).
    .PARAMETER OnProgress
        { param() } after each page or error.
    .PARAMETER OnIdle
        { param([string]$Mailbox, [int]$Free) } when a mailbox has request slots that nothing waits for (its last lists
        are read page after page): returns new jobs for it (a slice of a list in course cut in two), or nothing.
    .NOTES
        Stopped by the user (window): OperationCanceledException at once; the requests in flight are abandoned.
    #>
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Jobs,
        [Parameter(Mandatory = $true)][scriptblock]$OnPage,
        [Parameter(Mandatory = $true)][scriptblock]$OnDone,
        [scriptblock]$OnProgress,
        [scriptblock]$OnIdle
    )

    if (-not $Jobs.Count) { return }
    $lists = $Jobs.Count
    # Time of each page (request sent -> answer read): average, longest, and the requests in flight on average.
    $pageSeconds = [double]0; $pageMax = [double]0; $answers = 0
    $settings = $script:Graph.Settings
    $maxRetries = [int]$settings.MaxRetries
    $concurrency = [Math]::Max(1, [int]$settings.MaxConcurrency)
    $limit = $script:MailboxConcurrency
    # One queue per mailbox; 'ready' holds the mailboxes that have a job waiting and a free slot, each once.
    $queues = [Collections.Generic.Dictionary[string, Collections.Generic.LinkedList[object]]]::new([StringComparer]::OrdinalIgnoreCase)
    $busy = [Collections.Generic.Dictionary[string, int]]::new([StringComparer]::OrdinalIgnoreCase)
    $ready = [Collections.Generic.Queue[string]]::new()
    $inReady = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $delayed = [Collections.Generic.List[object]]::new()
    $inflight = [Collections.Generic.List[object]]::new()
    $markReady = {
        param([string]$Key)
        $q = $queues[$Key]
        if ($q.Count -and $busy[$Key] -lt $limit -and $inReady.Add($Key)) { $ready.Enqueue($Key) }
    }
    foreach ($job in $Jobs) {
        if (-not $queues.ContainsKey($job.Mailbox)) { $queues[$job.Mailbox] = [Collections.Generic.LinkedList[object]]::new(); $busy[$job.Mailbox] = 0 }
        [void]$queues[$job.Mailbox].AddLast($job)
    }
    foreach ($key in @($queues.Keys)) { & $markReady $key }
    $pages = 0; $retries = 0; $started = [datetime]::UtcNow

    while ($ready.Count -or $delayed.Count -or $inflight.Count) {
        if ($script:Ui -and $script:Ui.Cancel) { throw [OperationCanceledException]::new('Stopped by the user.') }

        # ---- retries whose delay is over go back first in the queue of their mailbox -------------------------
        if ($delayed.Count) {
            $now = [datetime]::UtcNow
            foreach ($job in @($delayed | Where-Object { $_.NotBefore -le $now })) {
                [void]$delayed.Remove($job)
                [void]$queues[$job.Mailbox].AddFirst($job)
                & $markReady $job.Mailbox
            }
        }

        # ---- new requests ------------------------------------------------------------------------------------
        while ($inflight.Count -lt $concurrency -and $ready.Count) {
            $key = $ready.Dequeue()
            [void]$inReady.Remove($key)
            $q = $queues[$key]
            if (-not $q.Count -or $busy[$key] -ge $limit) { continue }
            $job = $q.First.Value
            $q.RemoveFirst()
            $busy[$key] = $busy[$key] + 1
            Update-MmrToken
            if (-not $job.Started) { $job.Started = [datetime]::UtcNow }
            $inflight.Add(@{ Job = $job; Sent = [datetime]::UtcNow; Handle = (Start-MmrGraphSend -Method 'GET' -Url $job.Url -Headers $job.Headers) })
            # The mailbox in turn again: the others are served before its next job.
            & $markReady $key
        }

        if (-not $inflight.Count) {
            if ($delayed.Count) {
                $wait = ($delayed | ForEach-Object { $_.NotBefore } | Measure-Object -Minimum).Minimum
                Wait-MmrUi ([int][Math]::Max(20, [Math]::Min(5000, ($wait - [datetime]::UtcNow).TotalMilliseconds))) -NoCancel
            }
            continue
        }

        # ---- answers -------------------------------------------------------------------------------------------
        $finished = @($inflight | Where-Object { Test-MmrGraphSendDone $_.Handle })
        if (-not $finished.Count) {
            $tasks = @($inflight | ForEach-Object { $_.Handle.Task } | Where-Object { $_ })
            if ($tasks.Count -and -not ($script:Ui -and $script:Ui.Pump)) { [void][Threading.Tasks.Task]::WaitAny([Threading.Tasks.Task[]]$tasks, 50) }
            else { Wait-MmrUi 20 -NoCancel }
            continue
        }
        foreach ($call in $finished) {
            [void]$inflight.Remove($call)
            $job = $call.Job
            $busy[$job.Mailbox] = $busy[$job.Mailbox] - 1
            $r = Complete-MmrGraphSend $call.Handle
            $took = ([datetime]::UtcNow - $call.Sent).TotalSeconds
            $job.Seconds += $took; $pageSeconds += $took; $answers++
            if ($took -gt $pageMax) { $pageMax = $took }
            if ($r.Status -eq 200) {
                # A page whose JSON is cut short is asked again, smaller (Step-MmrPagedCut): never the end of the run.
                $next = $null
                $cut = $null
                try { $next = [string](& $OnPage $job $r.Content) }
                catch { $cut = Get-MmrCutAnswer $_.Exception; if (-not $cut) { throw } }
                if ($cut) {
                    $retries++
                    $step = Step-MmrPagedCut -Job $job -Length $r.Content.Length -Reason $cut -MaxRetries $maxRetries
                    if ($step.Retry) { $job.NotBefore = [datetime]::UtcNow.AddSeconds($step.Delay); $delayed.Add($job) }
                    else { & $OnDone $job 502 'CutAnswer' $step.Message }
                }
                else {
                    $job.Attempt = 0
                    $job.Pages++
                    $pages++
                    if ($job.CutTotal) { $next = Get-MmrPagedNext -Job $job -Next $next }
                    if ($next) {
                        $job.Url = $next
                        [void]$queues[$job.Mailbox].AddFirst($job)
                    }
                    else { & $OnDone $job 200 '' '' }
                }
            }
            elseif (($r.Status -eq 401 -and $job.Attempt -lt 1) -or ($r.Status -in 0, 429, 500, 502, 503, 504 -and $job.Attempt -lt $maxRetries)) {
                if ($r.Status -eq 401) { Update-MmrToken -Force; $delay = 0 }
                else { $delay = if ($r.RetryAfter -gt 0) { $r.RetryAfter } else { [Math]::Min(60, [Math]::Pow(2, $job.Attempt + 1)) } }
                $job.Attempt++
                $retries++
                $job.NotBefore = [datetime]::UtcNow.AddSeconds($delay)
                $delayed.Add($job)
                Write-MmrLog 'WARN' ("Graph GET $($job.Url) -> $($r.Status)$(if ($r['Error']) { " ($($r['Error']))" }), retry $($job.Attempt) in $delay s")
            }
            else {
                $parsed = ConvertFrom-MmrJson $r.Content
                $err = Get-MmrGraphError $parsed
                if ($r.Status -eq 0) { $err = @{ Code = 'NoResponse'; Message = [string]$r['Error'] } }
                & $OnDone $job $r.Status $err.Code $err.Message
            }
            & $markReady $job.Mailbox
            # Slots of this mailbox that nothing waits for: the caller may cut a list in course in two.
            if ($OnIdle) {
                $key = $job.Mailbox
                $free = $limit - $busy[$key] - $queues[$key].Count
                if ($free -gt 0) {
                    foreach ($new in @(& $OnIdle $key $free)) {
                        if (-not $new) { continue }
                        if (-not $queues.ContainsKey($new.Mailbox)) { $queues[$new.Mailbox] = [Collections.Generic.LinkedList[object]]::new(); $busy[$new.Mailbox] = 0 }
                        [void]$queues[$new.Mailbox].AddLast($new)
                        $lists++
                        & $markReady $new.Mailbox
                    }
                }
            }
        }
        if ($OnProgress) { & $OnProgress }
    }
    $elapsed = ([datetime]::UtcNow - $started).TotalSeconds
    $speed = if ($answers) { '; a page {0:N1} s on average, {1:N1} s at most, {2:N1} requests at a time' -f ($pageSeconds / $answers), $pageMax, ($pageSeconds / [Math]::Max(0.001, $elapsed)) } else { '' }
    Write-MmrLog 'INFO' ("Graph: {0} list(s), {1} page(s), {2} retried, {3}{4}" -f $lists, $pages, $retries, (Format-MmrDuration $elapsed), $speed)
}
