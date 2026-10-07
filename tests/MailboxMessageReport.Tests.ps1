#Requires -Version 7.4
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '6.1.0' }
<#
    Mailbox Message Report - automated tests (Pester 6.1 or later).
    Author  : Nicolas Fabert
    Version : 1.0.0

    Run:  .\Run-Tests.ps1      (or Invoke-Pester -Path .\tests -Output Detailed)

    No tenant is needed: Start-MmrGraphSend is replaced by a simulated Exchange Online tenant
    (tests\MailboxMessageReport.FakeGraph.ps1) that answers like Graph did in the lab: the archive ID from
    settings/exchange (beta), the archive opened by its MBX ID, the folder tree from mailFolders/delta, the messages
    filtered by date and subject (contains() refused before the date, as Exchange does), page after page.
#>

BeforeAll {
    $script:Root = Split-Path $PSScriptRoot -Parent
    Import-Module (Join-Path $script:Root 'MailboxMessageReport.psd1') -Force
    $script:Module = Get-Module MailboxMessageReport
    . (Join-Path $PSScriptRoot 'MailboxMessageReport.FakeGraph.ps1')
    & $script:Module { $script:Quiet = $true }

    $script:Tenant = '11111111-2222-3333-4444-555555555555'
    function New-TestSettings([hashtable]$Overrides = @{}) {
        $s = & $script:Module { Get-MmrDefaultConfiguration }
        $s.TenantId = $script:Tenant; $s.AppId = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'; $s.CertificateThumbprint = ('AB' * 20)
        $s.TimeZone = 'UTC'; $s.MaxRetries = 3; $s.MaxConcurrency = 8
        $s.OutputPath = Join-Path $script:Root 'artifacts\test-reports'; $s.LogPath = Join-Path $script:Root 'artifacts\test-logs'
        foreach ($k in $Overrides.Keys) { $s[$k] = $Overrides[$k] }
        return $s
    }
    function Connect-Test([hashtable]$Settings, [string[]]$Roles) {
        $token = if ($Roles) { New-FakeToken -TenantId $script:Tenant -Roles $Roles } else { New-FakeToken -TenantId $script:Tenant }
        Mock -ModuleName MailboxMessageReport Get-MmrCertificate { [pscustomobject]@{ Thumbprint = 'AB'; NotAfter = (Get-Date).AddYears(1) } }
        Mock -ModuleName MailboxMessageReport Get-MmrAppToken { @{ Token = $token; ExpiresUtc = [datetime]::UtcNow.AddHours(1) } }.GetNewClosure()
        Connect-MmrGraph -Settings $Settings
    }
    # A search of the simulated tenant, from the request to the result (steps 2 to 4), as the command line does.
    function Find-Test([string[]]$Mailbox, [hashtable]$More = @{}, [hashtable]$Overrides = @{}, [string[]]$Roles) {
        $s = New-TestSettings $Overrides
        $null = Connect-Test $s $Roles
        $a = @{ Settings = $s; Mailbox = $Mailbox }
        foreach ($k in $More.Keys) { $a[$k] = $More[$k] }
        $request = New-MmrRequest @a
        $run = New-MmrRunFolder -OutputPath $s.OutputPath -Prefix $s.ReportPrefix
        & $script:Module { Initialize-MmrSteps -Total 5 }
        $result = Find-MmrMessages -Settings $s -Request $request -PartsPath (Join-Path $run '.parts')
        [pscustomobject]@{ Result = $result; Run = $run; Request = $request; Settings = $s }
    }
    function Export-Test($Search, [string]$Layout = 'Global', [string[]]$Formats = @('Csv', 'Html'), [int]$HtmlMax = 20000, [int]$Preview = 0) {
        Export-MmrReport -Result $Search.Result -Directory $Search.Run -Prefix 'MailboxMessageReport' -Formats $Formats -Layout $Layout -HtmlMaxMessages $HtmlMax -PreviewMessages $Preview -PartsPath (Join-Path $Search.Run '.parts')
    }
    function Read-TestCsv([string]$Path) { Import-Csv -LiteralPath $Path -Delimiter ';' }

    # Every Graph request goes to the simulated tenant; the requests in flight are counted per mailbox.
    Mock -ModuleName MailboxMessageReport Start-MmrGraphSend {
        $key = & (Get-Module MailboxMessageReport) { param($u) Get-MmrMailboxKey ($u -replace '^https://graph\.microsoft\.com/(v1\.0|beta)', '') } $Url
        $script:Fake.Open[$key] = [int]$script:Fake.Open[$key] + 1
        if ($script:Fake.Open[$key] -gt [int]$script:Fake.MaxOpen[$key]) { $script:Fake.MaxOpen[$key] = $script:Fake.Open[$key] }
        [pscustomobject]@{ Task = $null; Request = $null; Key = $key; Response = (Invoke-FakeGraphHttp -Method $Method -Url $Url -Body $Body -Headers $Headers) }
    }
    Mock -ModuleName MailboxMessageReport Complete-MmrGraphSend { $script:Fake.Open[$Handle.Key] = [int]$script:Fake.Open[$Handle.Key] - 1; $Handle.Response }

    # The tenant of most tests: Alice (archive, alias), Bob (no archive), Carol (on-premises).
    function New-TestTenant {
        Reset-FakeTenant -TenantId $script:Tenant
        $script:Alice = Add-FakeUser 'alice@contoso.test' -Name 'Alice Archive' -Aliases 'alice.alias@contoso.test' -Archive
        $script:Bob = Add-FakeUser 'bob@contoso.test' -Name 'Bob Primary'
        $script:Carol = Add-FakeUser 'carol@contoso.test' -Name 'Carol OnPrem' -OnPremises
        $null = Add-FakeFolder 'alice@contoso.test' -Path '\Inbox'
        $null = Add-FakeFolder 'alice@contoso.test' -Path '\Sent Items'
        $null = Add-FakeMessage 'alice@contoso.test' -Path '\Inbox' -Subject 'Contrat Alpha - v3' -Received '2026-03-10T09:00:00' -To 'alice@contoso.test' -Cc 'dan@contoso.test' -Attachments
        $null = Add-FakeMessage 'alice@contoso.test' -Path '\Inbox' -Subject 'Weekly status' -Received '2026-04-01T08:00:00' -To 'alice@contoso.test'
        $null = Add-FakeMessage 'alice@contoso.test' -Path '\Inbox\Projects\Alpha' -Subject 'Projet Alpha: kick-off' -Received '2025-11-05T14:00:00' -To 'alice@contoso.test', 'bob@contoso.test'
        $null = Add-FakeMessage 'alice@contoso.test' -Path '\Sent Items' -Subject 'RE: Contrat Alpha' -Received '2026-03-11T10:00:00' -From 'alice@contoso.test' -To 'ext@fabrikam.test' -Bcc 'boss@contoso.test'
        $null = Add-FakeMessage 'alice@contoso.test' -Location Archive -Path '\Inbox' -Subject 'Contrat Alpha - v1' -Received '2019-02-01T09:00:00' -To 'alice@contoso.test'
        $null = Add-FakeMessage 'alice@contoso.test' -Location Archive -Path '\2019\Q1' -Subject 'Facture 2019-001' -Received '2019-01-15T09:00:00' -To 'alice@contoso.test'
        $null = Add-FakeMessage 'alice@contoso.test' -Location Archive -Path '\2019\Q1' -Subject '=SUM(A1:A9) <script>alert(1)</script>' -Received '2019-03-20T09:00:00' -To 'alice@contoso.test'
        $null = Add-FakeMessage 'alice@contoso.test' -Location Archive -Path '\Junk Email' -Subject 'Contrat Alpha spam' -Received '2019-05-01T09:00:00' -To 'alice@contoso.test'
        $null = Add-FakeFolder 'alice@contoso.test' -Location Archive -Path '\Empty'
        $null = Add-FakeMessage 'alice@contoso.test' -Path '\Deletions' -Recoverable -Subject 'Contrat Alpha deleted' -Received '2024-01-10T09:00:00' -To 'alice@contoso.test'
        $null = Add-FakeFolder 'alice@contoso.test' -Path '\Purges' -Recoverable
        $null = Add-FakeMessage 'alice@contoso.test' -Location Archive -Path '\Deletions' -Recoverable -Subject 'Old invoice deleted' -Received '2018-06-10T09:00:00' -To 'alice@contoso.test'
        $null = Add-FakeMessage 'bob@contoso.test' -Path '\Inbox' -Subject 'Hello Bob' -Received '2026-01-02T09:00:00' -To 'bob@contoso.test' -Type '#microsoft.graph.eventMessageRequest'
        $null = Add-FakeMessage 'bob@contoso.test' -Path '\Inbox' -Subject "L'offre Contrat Alpha" -Received '2026-02-02T09:00:00' -To 'bob@contoso.test'
    }
}

Describe 'Configuration and request' {
    It 'reads the delivered configuration and resolves the paths from the tool folder' {
        $c = Import-MmrConfiguration
        $c.OutputPath | Should -Be (Join-Path $script:Root 'reports')
        $c.LogPath | Should -Be (Join-Path $script:Root 'logs')
        $c.Locations | Should -Be @('Primary', 'Archive')
        $c.RecoverableItems | Should -BeFalse
        $c.PageSize | Should -Be 250
        $c.SplitFolderItems | Should -Be 5000
        $c.Recipients | Should -BeTrue
        $c.ReportLayout | Should -Be 'Global'
        $c.HtmlMaxMessages | Should -Be 20000
    }

    It 'lists unknown sections, unknown keys and invalid values together' {
        $path = Join-Path $script:Root 'artifacts\bad.config.psd1'
        [void][IO.Directory]::CreateDirectory((Split-Path $path))
        "@{ Tenant = @{ TenantId = 'not a tenant' }; Search = @{ Locations = @('Primary', 'Elsewhere'); Futur = 1; ExcludeFolders = @('Junk') }; Extra = @{}; Graph = @{ PageSize = 5000 }; Report = @{ CsvDelimiter = '|'; Layout = 'Flat' } }" | Set-Content $path
        $text = ({ Import-MmrConfiguration -Path $path } | Should -Throw -PassThru).Exception.Message
        $text | Should -Match "Unknown key 'Search.Futur'"
        $text | Should -Match "Unknown section 'Extra'"
        $text | Should -Match 'Tenant.TenantId must be'
        $text | Should -Match 'Search.Locations must contain'
        $text | Should -Match 'Search.ExcludeFolders'
        $text | Should -Match 'Graph.PageSize must be'
        $text | Should -Match 'Report.CsvDelimiter must be'
        $text | Should -Match 'Report.Layout must be'
    }

    It 'reads a list of mailboxes: text, or CSV with the ArchiveGuid of Get-EXOMailbox' {
        $dir = Join-Path $script:Root 'artifacts'
        $txt = Join-Path $dir 'list.txt'
        "# team`nalice@contoso.test`nBob@Contoso.test  # comment`n`nalice@contoso.test" | Set-Content $txt
        $l = Read-MmrMailboxFile $txt
        ($l | ForEach-Object Address) | Should -Be @('alice@contoso.test', 'bob@contoso.test')
        $csv = Join-Path $dir 'list.csv'
        "`"PrimarySmtpAddress`",`"ArchiveGuid`"`n`"alice@contoso.test`",`"{01234567-89AB-CDEF-0123-456789ABCDEF}`"`n`"bob@contoso.test`",`"00000000-0000-0000-0000-000000000000`"" | Set-Content $csv
        $l = Read-MmrMailboxFile $csv
        $l[0].ArchiveGuid | Should -Be '01234567-89ab-cdef-0123-456789abcdef'
        $l[1].ArchiveGuid | Should -Be ''
    }

    It 'builds a request: mailboxes typed and of a file once each, the period in the time zone, the subjects, the defaults' {
        $s = New-TestSettings @{ TimeZone = 'Romance Standard Time' }
        $csv = Join-Path $script:Root 'artifacts\list2.csv'
        "Mail;ArchiveGuid`nbob@contoso.test;`ncarol@contoso.test;01234567-89ab-cdef-0123-456789abcdef" | Set-Content $csv
        $r = New-MmrRequest -Settings $s -Mailbox 'alice@contoso.test; Bob@contoso.test' -MailboxFile $csv -Start '2026-01-01' -End '2026-01-31' -Subject "Contrat Alpha`nFacture", 'Facture'
        ($r.Mailboxes | ForEach-Object Address) | Should -Be @('alice@contoso.test', 'bob@contoso.test', 'carol@contoso.test')
        ($r.Mailboxes | Where-Object Address -eq 'carol@contoso.test').ArchiveGuid | Should -Be '01234567-89ab-cdef-0123-456789abcdef'
        $r.Start | Should -Be ([datetime]'2025-12-31T23:00:00')
        $r.End | Should -Be ([datetime]'2026-01-31T23:00:00')
        $r.Subject | Should -Be @('Contrat Alpha', 'Facture')
        $r.Location | Should -Be @('Primary', 'Archive')
        $r.Layout | Should -Be 'Global'
        (Test-MmrRequest $r).IsValid | Should -BeTrue
        $none = New-MmrRequest -Settings $s -Mailbox 'a@contoso.test'
        $none.Start | Should -BeNullOrEmpty
        $none.End | Should -BeNullOrEmpty
        $past = New-MmrRequest -Settings (New-TestSettings @{ PastDays = 30 }) -Mailbox 'a@contoso.test'
        $past.Start | Should -Be ([datetime]::UtcNow.Date.AddDays(-30))
    }

    It 'lists every problem of a request' {
        $s = New-TestSettings
        $r = New-MmrRequest -Settings $s -Mailbox 'not-an-address' -Start '2026-02-01' -End '2026-01-01' -Location 'Primary' -Layout 'Global'
        $r.Location = @('Elsewhere'); $r.Layout = 'Flat'; $r.Mailboxes[0].ArchiveGuid = 'xyz'; $r.ExcludeFolders = @('Junk')
        $p = (Test-MmrRequest $r).Problems -join "`n"
        $p | Should -Match "'not-an-address' is not an SMTP address"
        $p | Should -Match "ArchiveGuid 'xyz' is not a GUID"
        $p | Should -Match 'end of the period must be after'
        $p | Should -Match "Location: 'Primary', 'Archive'"
        $p | Should -Match 'Report layout'
        $p | Should -Match "Folder left out 'Junk'"
        (Test-MmrRequest (New-MmrRequest -Settings $s)).Problems | Should -Match 'Give the mailbox'
    }

    It 'writes the filter Exchange accepts: the date first, then the subjects' {
        & $script:Module {
            Get-MmrMessageFilter | Should -Be ''
            Get-MmrMessageFilter -Subject 'Alpha' | Should -Be "receivedDateTime ge 1900-01-01T00:00:00Z and contains(subject,'Alpha')"
            Get-MmrMessageFilter -Start ([datetime]'2026-01-01T00:00:00') -End ([datetime]'2026-02-01T00:00:00') | Should -Be 'receivedDateTime ge 2026-01-01T00:00:00Z and receivedDateTime lt 2026-02-01T00:00:00Z'
            Get-MmrMessageFilter -End ([datetime]'2026-02-01T00:00:00') -Subject "L'offre", 'B' | Should -Be "receivedDateTime ge 1900-01-01T00:00:00Z and receivedDateTime lt 2026-02-01T00:00:00Z and (contains(subject,'L''offre') or contains(subject,'B'))"
        }
    }
}

Describe 'Microsoft Graph connection' {
    It 'needs Mail.ReadBasic.All (or Mail.Read) and says how to grant it' {
        $s = New-TestSettings
        { Connect-Test $s @('User.Read.All') } | Should -Throw '*Mail.ReadBasic.All*Grant admin consent*'
        $c = Connect-Test $s @('Mail.ReadBasic.All')
        $c.CanReadMail | Should -BeTrue
        $c.CanReadUsers | Should -BeFalse
        $c.CanFindUsers | Should -BeFalse
        $c = Connect-Test $s @('Mail.Read', 'User.ReadBasic.All')
        $c.CanFindUsers | Should -BeTrue
        $c.CanReadUsers | Should -BeFalse
    }

    It 'stops when the token belongs to another tenant' {
        $s = New-TestSettings
        $other = New-FakeToken -TenantId '99999999-2222-3333-4444-555555555555'
        Mock -ModuleName MailboxMessageReport Get-MmrCertificate { [pscustomobject]@{ Thumbprint = 'AB'; NotAfter = (Get-Date).AddYears(1) } }
        Mock -ModuleName MailboxMessageReport Get-MmrAppToken { @{ Token = $other; ExpiresUtc = [datetime]::UtcNow.AddHours(1) } }.GetNewClosure()
        { Connect-MmrGraph -Settings $s } | Should -Throw '*Nothing was read*'
    }
}

Describe 'Mailboxes' {
    BeforeEach { New-TestTenant }

    It 'finds the user by any alias and its archive in Graph, without Exchange Online PowerShell' {
        $null = Connect-Test (New-TestSettings)
        $m = Resolve-MmrMailboxes -Entries @([pscustomobject]@{ Address = 'alice.alias@contoso.test'; ArchiveGuid = '' }, [pscustomobject]@{ Address = 'bob@contoso.test'; ArchiveGuid = '' })
        $m[0].Address | Should -Be 'alice@contoso.test'
        $m[0].DisplayName | Should -Be 'Alice Archive'
        $m[0].Archive | Should -Be 'Yes'
        $m[0].ArchiveSource | Should -Be 'Graph'
        $m[0].ArchiveKey | Should -Be $script:Alice.ArchiveKey
        $m[0].ArchiveGuid | Should -Be $script:Alice.ArchiveGuid
        $m[1].Archive | Should -Be 'No'
        @($script:Fake.Calls | Where-Object { $_.Version -eq 'beta' -and $_.Url -like '*/settings/exchange' }).Count | Should -Be 2
    }

    It 'says why a mailbox cannot be read: on-premises or inactive, not a mailbox' {
        $r = (Find-Test 'carol@contoso.test', 'nobody@contoso.test', 'bob@contoso.test').Result
        $carol = $r.Mailboxes | Where-Object Address -eq 'carol@contoso.test'
        $carol.State | Should -Be 'NotReachable'
        $carol.Detail | Should -Match 'on-premises'
        $carol.Status | Should -Be 'Not read'
        $nobody = $r.Mailboxes | Where-Object Address -eq 'nobody@contoso.test'
        $nobody.State | Should -Be 'NotFound'
        $nobody.Notes | Should -Be ''
        $r.Status | Should -Be 'Warning'
        $r.Counts.MailboxesNotRead | Should -Be 2
        ($r.Mailboxes | Where-Object Address -eq 'bob@contoso.test').Messages | Should -Be 2
    }

    It 'reads the archive from the ArchiveGuid of the list without User.Read.All' {
        $guid = $script:Alice.ArchiveGuid
        $csv = Join-Path $script:Root 'artifacts\guid.csv'
        "PrimarySmtpAddress,ArchiveGuid`nalice@contoso.test,$guid`nbob@contoso.test," | Set-Content $csv
        $r = (Find-Test -Mailbox @() -More @{ MailboxFile = $csv } -Roles @('Mail.ReadBasic.All')).Result
        $alice = $r.Mailboxes | Where-Object Address -eq 'alice@contoso.test'
        $alice.ArchiveSource | Should -Be 'File'
        $alice.ArchiveMessages | Should -Be 4
        ($r.Mailboxes | Where-Object Address -eq 'bob@contoso.test').Archive | Should -Be 'Unknown'
        $r.Warnings -join ' ' | Should -Match 'No User.Read.All'
        @($script:Fake.Calls | Where-Object Url -match '^/users\?').Count | Should -Be 0
    }

    It 'reads a mailbox typed twice under two aliases once' {
        $r = (Find-Test 'alice@contoso.test', 'alice.alias@contoso.test').Result
        @($r.Mailboxes).Count | Should -Be 1
    }
}

Describe 'Folders and messages' {
    BeforeEach { New-TestTenant }

    It 'reads every folder of the primary mailbox and of the archive, with its path; empty folders are not read' {
        $s = Find-Test 'alice@contoso.test'
        $f = $s.Result.Folders
        ($f | Where-Object Location -eq 'Primary').Path | Should -Be @('\Inbox', '\Inbox\Projects', '\Inbox\Projects\Alpha', '\Sent Items')
        ($f | Where-Object Location -eq 'Archive').Path | Should -Be @('\2019', '\2019\Q1', '\Empty', '\Inbox', '\Junk Email')
        ($f | Where-Object Path -eq '\Inbox\Projects').Status | Should -Be 'Empty'
        @($script:Fake.Calls | Where-Object { $_.Url -like '*/messages?*' }).Count | Should -Be 6
        $s.Result.Counts.Messages | Should -Be 8
        $s.Result.Counts.PrimaryMessages | Should -Be 4
        $s.Result.Counts.ArchiveMessages | Should -Be 4
        $s.Result.Status | Should -Be 'Completed'
        # The folder tree in pages of 200.
        @($script:Fake.Calls | Where-Object { $_.Url -like '*/mailFolders/delta*' }).Count | Should -Be 2
    }

    It 'reads Recoverable Items when asked, primary mailbox and archive' {
        $r = (Find-Test 'alice@contoso.test' -More @{ RecoverableItems = $true }).Result
        $ri = @($r.Folders | Where-Object RecoverableItems)
        $ri.Path | Should -Contain '\Recoverable Items\Deletions'
        $ri.Path | Should -Contain '\Recoverable Items\Purges'
        $r.Counts.RecoverableMessages | Should -Be 2
        $r.Counts.Messages | Should -Be 10
    }

    It 'keeps the messages of the period and of the subjects; leaves out the folders asked' {
        $s = Find-Test 'alice@contoso.test', 'bob@contoso.test' -More @{ Subject = 'contrat alpha'; Start = [datetime]'2019-01-01'; End = [datetime]'2026-03-10'; ExcludeFolder = '\Junk*'; RecoverableItems = $true }
        $r = $s.Result
        ($r.Folders | Where-Object { $_.Location -eq 'Archive' -and $_.Path -eq '\Junk Email' }).Status | Should -Be 'Excluded'
        $report = Export-Test $s
        $rows = @(Read-TestCsv $report.Files.Messages)
        # Alice: primary mailbox, its Recoverable Items, archive; then Bob.
        $rows.Subject | Should -Be @('Contrat Alpha - v3', 'Contrat Alpha deleted', 'Contrat Alpha - v1', "L'offre Contrat Alpha")
        $rows[0].Location | Should -Be 'Primary'
        ($rows | Where-Object Subject -eq 'Contrat Alpha deleted').RecoverableItems | Should -Be 'Yes'
        ($rows | Where-Object Subject -eq 'Contrat Alpha - v1').Location | Should -Be 'Archive'
        # Exchange refuses contains() before the date: the filter of every request starts with the date.
        @($script:Fake.Calls | Where-Object { $_.Url -like '*/messages?*' -and $_.Url -notlike "*`$filter=receivedDateTime ge 2019-01-01T00:00:00Z and receivedDateTime lt 2026-03-11T00:00:00Z and contains(subject,'contrat alpha')*" }).Count | Should -Be 0
    }

    It 'writes the columns of a message: dates in the time zone, sender, recipients, message ID, folder, location' {
        $s = Find-Test 'alice@contoso.test' -Overrides @{ TimeZone = 'Romance Standard Time' }
        $rows = @(Read-TestCsv (Export-Test $s).Files.Messages)
        $m = $rows | Where-Object Subject -eq 'Contrat Alpha - v3'
        $m.Mailbox | Should -Be 'alice@contoso.test'
        $m.MailboxName | Should -Be 'Alice Archive'
        $m.FolderPath | Should -Be '\Inbox'
        $m.Folder | Should -Be 'Inbox'
        $m.Received | Should -Be '2026-03-10 10:00:00'
        $m.ReceivedUtc | Should -Be '2026-03-10T09:00:00Z'
        $m.Sent | Should -Be '2026-03-10 09:59:00'
        $m.From | Should -Be 'sender@fabrikam.test'
        $m.To | Should -Be 'alice@contoso.test'
        $m.Cc | Should -Be 'dan@contoso.test'
        $m.HasAttachments | Should -Be 'Yes'
        $m.InternetMessageId | Should -Match '^<.+@fabrikam.test>$'
        $m.Type | Should -Be 'Message'
        ($rows | Where-Object Subject -eq 'RE: Contrat Alpha').Bcc | Should -Be 'boss@contoso.test'
        ($rows | Where-Object Subject -eq 'Projet Alpha: kick-off').To | Should -Be 'alice@contoso.test; bob@contoso.test'
        # The order of the report: primary mailbox then archive, folder path, newest first.
        $rows.Subject | Should -Be @('Weekly status', 'Contrat Alpha - v3', 'Projet Alpha: kick-off', 'RE: Contrat Alpha', "'=SUM(A1:A9) <script>alert(1)</script>", 'Facture 2019-001', 'Contrat Alpha - v1', 'Contrat Alpha spam')
        $bob = Find-Test 'bob@contoso.test'
        (@(Read-TestCsv (Export-Test $bob).Files.Messages) | Where-Object Subject -eq 'Hello Bob').Type | Should -Be 'Meeting request'
    }

    It 'follows the pages of a large folder and reads at most 4 lists at a time per mailbox' {
        Reset-FakeTenant -TenantId $script:Tenant
        $null = Add-FakeUser 'big@contoso.test' -Archive
        foreach ($i in 1..10) { foreach ($j in 1..13) { $null = Add-FakeMessage 'big@contoso.test' -Location Archive -Path "\Year $i" -Subject "Message $i-$j" -Received ([datetime]'2020-01-01').AddDays($i * 20 + $j) } }
        foreach ($j in 1..25) { $null = Add-FakeMessage 'big@contoso.test' -Path '\Inbox' -Subject "Inbox $j" -Received ([datetime]'2026-01-01').AddHours($j) }
        $r = (Find-Test 'big@contoso.test' -Overrides @{ PageSize = 10 }).Result
        $r.Counts.Messages | Should -Be 155
        ($r.Folders | Where-Object Path -eq '\Inbox').Pages | Should -Be 3
        $script:Fake.MaxOpen[$script:Fake.Users[0].ArchiveKey.ToLowerInvariant()] | Should -Be 4
        $script:Fake.MaxOpen.Values | ForEach-Object { $_ | Should -BeLessOrEqual 4 }
    }

    It 'reads a large folder in slices of its dates, 4 at a time, every message once, newest first' {
        Reset-FakeTenant -TenantId $script:Tenant
        $null = Add-FakeUser 'big@contoso.test' -Archive
        foreach ($j in 1..130) { $null = Add-FakeMessage 'big@contoso.test' -Location Archive -Path '\Inbox' -Subject "Message $j" -Received ([datetime]'2015-01-01').AddDays($j * 11).AddSeconds($j) }
        $null = Add-FakeMessage 'big@contoso.test' -Path '\Inbox' -Subject 'Small' -Received '2026-01-01'
        $s = Find-Test 'big@contoso.test' -Overrides @{ SplitFolderItems = 30; PageSize = 10 }
        $f = $s.Result.Folders | Where-Object { $_.Location -eq 'Archive' -and $_.Path -eq '\Inbox' }
        $f.Slices | Should -Be 5
        $f.Messages | Should -Be 130
        $f.Detail | Should -Be 'read in 5 slices of its dates'
        ($s.Result.Folders | Where-Object Location -eq 'Primary').Slices | Should -Be 1
        $script:Fake.MaxOpen[$script:Fake.Users[0].ArchiveKey.ToLowerInvariant()] | Should -Be 4
        $rows = @(Read-TestCsv (Export-Test $s).Files.Messages | Where-Object Location -eq 'Archive')
        $rows.Count | Should -Be 130
        @($rows.Subject | Select-Object -Unique).Count | Should -Be 130
        $dates = @($rows | ForEach-Object ReceivedUtc)
        ($dates -join ',') | Should -Be (($dates | Sort-Object -Descending) -join ',')
        # With a subject: the slices keep it, after their dates.
        $sub = (Find-Test 'big@contoso.test' -More @{ Subject = 'Message 1' } -Overrides @{ SplitFolderItems = 30 }).Result
        ($sub.Folders | Where-Object { $_.Location -eq 'Archive' -and $_.Path -eq '\Inbox' }).Messages | Should -Be 42   # Message 1, 10 to 19, 100 to 130
    }

    It 'reads the messages without their recipients when asked' {
        $s = Find-Test 'alice@contoso.test' -More @{ Recipients = $false }
        @($script:Fake.Calls | Where-Object { $_.Url -like '*/messages?*' -and $_.Url -match 'Recipients' }).Count | Should -Be 0
        $rows = @(Read-TestCsv (Export-Test $s).Files.Messages)
        $rows.Count | Should -Be 8
        @($rows | Where-Object { $_.To -or $_.Cc -or $_.Bcc }).Count | Should -Be 0
        ($rows | Where-Object Subject -eq 'Contrat Alpha - v3').From | Should -Be 'sender@fabrikam.test'
    }

    It 'retries a page throttled (429) and reports a folder that fails, the others are read' {
        $script:Fake.Throttle['*/mailFolders/*/messages*'] = 3
        $archiveInbox = (Get-FakeStore 'alice@contoso.test' Archive).Folders | Where-Object displayName -eq 'Inbox'
        $script:Fake.Fail["*$($archiveInbox.id)/messages*"] = @{ Status = 500; Code = 'ErrorInternalServerError' }
        $r = (Find-Test 'alice@contoso.test' -Overrides @{ MaxRetries = 2 }).Result
        $r.Counts.Messages | Should -Be 7
        $failed = $r.Folders | Where-Object Status -eq 'Failed'
        $failed.Path | Should -Be '\Inbox'
        $failed.Location | Should -Be 'Archive'
        $failed.Detail | Should -Match 'auxiliary archive'
        $r.Status | Should -Be 'Warning'
        ($r.Mailboxes[0]).Status | Should -Be 'Partial'
    }

    It 'stops at once when the window asks it' {
        $s = New-TestSettings
        $null = Connect-Test $s
        $request = New-MmrRequest -Settings $s -Mailbox 'alice@contoso.test'
        & $script:Module {
            param($s, $request)
            $script:Ui = @{ Cancel = $true; Queue = $null; Sink = $null; Pump = $null }
            try { { Find-MmrMessages -Settings $s -Request $request -PartsPath (Join-Path $s.OutputPath 'stopped\.parts') } | Should -Throw -ExceptionType ([OperationCanceledException]) }
            finally { $script:Ui = $null }
        } $s $request
    }
}

Describe 'Report' {
    BeforeEach { New-TestTenant }

    It 'writes one report for every mailbox: CSV safe for Excel, HTML safe, summary, the part files deleted' {
        $s = Find-Test 'alice@contoso.test', 'bob@contoso.test' -More @{ RecoverableItems = $true }
        $report = Export-Test $s -Preview 3
        $report.Files.Keys | Should -Be @('Messages', 'Mailboxes', 'Folders', 'Html', 'Summary')
        Test-Path (Join-Path $s.Run '.parts') | Should -BeFalse
        $raw = Get-Content $report.Files.Messages -Raw -Encoding utf8
        $raw | Should -Match ";'=SUM\(A1:A9\)"
        @(Read-TestCsv $report.Files.Messages).Count | Should -Be 12
        $html = Get-Content $report.Files.Html -Raw
        $html | Should -Not -Match '\{\{[A-Z_]+\}\}'
        $html | Should -Not -Match '<script>alert'
        $html | Should -Match 'Mailbox Message Report \| Alice Archive, Bob Primary'
        $summary = Get-Content $report.Files.Summary -Raw | ConvertFrom-Json
        $summary.Counts.Messages | Should -Be 12
        $summary.Counts.RecoverableMessages | Should -Be 2
        @($summary.Folders | Where-Object { $_.PSObject.Properties['Part'] }).Count | Should -Be 0
        $report.Preview.Lines.Count | Should -Be 3
        $report.Preview.Seen | Should -Be 12
        $rows = [MailboxMessageReportNative.PreviewRow]::Build($report.Preview)
        $rows[0].Location | Should -Be 'Primary'
        @(Read-TestCsv $report.Files.Mailboxes).Count | Should -Be 2
        (@(Read-TestCsv $report.Files.Folders) | Where-Object { $_.Path -eq '\Inbox\Projects\Alpha' }).Messages | Should -Be '1'
    }

    It 'keeps the first messages in the HTML report and says so; the CSV file holds them all' {
        $s = Find-Test 'alice@contoso.test'
        $report = Export-Test $s -HtmlMax 3
        $html = Get-Content $report.Files.Html -Raw
        $json = [regex]::Match($html, '<script type="application/json" id="data-summary">(.*?)</script>', 'Singleline').Groups[1].Value | ConvertFrom-Json
        $json.MessagesShown | Should -Be 3
        $json.MessagesTotal | Should -Be 8
        $json.Columns[0] | Should -Be 'Mailbox'
        @(Read-TestCsv $report.Files.Messages).Count | Should -Be 8
    }

    It 'writes one report per mailbox (read ones only), with a summary that links to them' {
        $s = Find-Test 'alice@contoso.test', 'bob@contoso.test', 'carol@contoso.test'
        $report = Export-Test $s -Layout PerMailbox
        $report.Files.Contains('Messages') | Should -BeFalse
        $report.MailboxFiles | Should -Be 2
        $dir = Join-Path $s.Run 'Mailboxes'
        (Get-ChildItem $dir).Name | Should -Be @('MailboxMessageReport-alice@contoso.test.csv', 'MailboxMessageReport-alice@contoso.test.html', 'MailboxMessageReport-bob@contoso.test.csv', 'MailboxMessageReport-bob@contoso.test.html')
        @(Read-TestCsv (Join-Path $dir 'MailboxMessageReport-bob@contoso.test.csv')).Count | Should -Be 2
        $html = Get-Content $report.Files.Html -Raw
        $html | Should -Match 'Mailboxes/MailboxMessageReport-alice@contoso.test.html'
        $json = [regex]::Match($html, 'id="data-summary">(.*?)</script>', 'Singleline').Groups[1].Value | ConvertFrom-Json
        $json.MessagesHere | Should -BeFalse
        $both = Export-Test (Find-Test 'alice@contoso.test') -Layout Both
        $both.Files.Contains('Messages') | Should -BeTrue
        $both.MailboxFiles | Should -Be 1
    }
}

Describe 'Window' {
    It 'builds the window from the configuration' {
        $s = New-TestSettings @{ Locations = @('Archive'); RecoverableItems = $true; ReportLayout = 'Both'; ExcludeFolders = @('\Junk Email') }
        $f = New-MmrForm -Configuration $s -Theme Light
        $c = $f.Controls
        $c.LocationPrimary.IsChecked | Should -BeFalse
        $c.LocationArchive.IsChecked | Should -BeTrue
        $c.RecoverableItems.IsChecked | Should -BeTrue
        $c.LayoutBoth.IsChecked | Should -BeTrue
        $c.ExcludeFolders.Text | Should -Be '\Junk Email'
        $c.TenantId.Text | Should -Be $script:Tenant
        $c.Search.IsEnabled | Should -BeTrue
        $c.LocationArchive.IsChecked = $false
        & $script:Module { Update-MmrGuiState }
        $c.Search.IsEnabled | Should -BeFalse
        $f.Form.Close()
    }

    It 'reads from the window: the mailboxes, a preview of the first messages, the report' {
        New-TestTenant
        $s = New-TestSettings @{ PreviewMessages = 5 }
        $null = Connect-Test $s
        $f = New-MmrForm -Configuration $s -Theme Light
        $c = $f.Controls
        $c.Mailbox.Text = "alice@contoso.test`r`nbob@contoso.test"
        $c.MailboxHint.Text | Should -Match '^2 mailboxes'
        try {
            & $script:Module { $script:GuiInline = $true; Invoke-MmrGuiSearch }
            $f.Mailboxes.Count | Should -Be 2
            $f.Mailboxes[0].GetType().Name | Should -Be 'MailboxRow'
            $f.Preview.Count | Should -Be 5
            $c.PreviewInfo.Text | Should -Match 'the first 5 of 10 messages'
            $c.Status.Text | Should -Match '^10 message'
            ($f.Lines -join "`n") | Should -Match 'Report: '
            & $script:Module { $script:Gui.LastReport } | Should -Match 'MailboxMessageReport\.html$'
            $c.OpenReport.IsEnabled | Should -BeTrue
        }
        finally { & $script:Module { $script:GuiInline = $false }; $f.Form.Close() }
    }

    It 'refuses to read with values to fix, and says which' {
        $s = New-TestSettings
        $s.AppId = ''
        $f = New-MmrForm -Configuration $s -Theme Light
        try {
            $f.Controls.Mailbox.Text = 'not an address'
            & $script:Module { Invoke-MmrGuiSearch }
            ($f.Lines -join "`n") | Should -Match "'not' is not an SMTP address"
            ($f.Lines -join "`n") | Should -Match 'Authentication.AppId is required'
            $f.Controls.Status.Text | Should -Be 'Fix the values on the left'
        }
        finally { $f.Form.Close() }
    }
}

Describe 'Progress' {
    It 'tells the time left from the speed of the progress, once it can be told' {
        & $script:Module {
            $script:ProgressEta = $null
            $t0 = [datetime]'2030-01-01T10:00:00Z'
            Get-MmrProgressEta -Fraction 0.10 -Text '18/180 folders read' -Now $t0 | Should -Be ''
            Get-MmrProgressEta -Fraction 0.30 -Text '54/180 folders read' -Now $t0.AddSeconds(4) | Should -Be 'about 15 s left'
            Get-MmrProgressEta -Fraction 1.00 -Text '180/180 folders read' -Now $t0.AddSeconds(13) | Should -Be ''
            $script:ProgressEta = $null
            Format-MmrTimeLeft 89 | Should -Be 'about 1 min 30 s left'
        }
    }
}

Describe 'Command line' {
    It 'stops with exit code 1 and the list of the problems on a wrong configuration' {
        $path = Join-Path $script:Root 'artifacts\cli-bad.config.psd1'
        "@{ Graph = @{ MaxConcurrency = 0 } }" | Set-Content $path
        $out = & pwsh -NoProfile -File (Join-Path $script:Root 'Invoke-MailboxMessageReport.ps1') -ConfigPath $path -Mailbox 'a@contoso.test' 2>&1
        $LASTEXITCODE | Should -Be 1
        ($out -join ' ') | Should -Match 'Graph.MaxConcurrency must be'
        Remove-Item $path
    }

    It 'refuses to run without a mailbox' {
        $path = Join-Path $script:Root 'artifacts\cli-ok.config.psd1'
        "@{ Logging = @{ Path = '$((Join-Path $script:Root 'artifacts\test-logs'))' } }" | Set-Content $path
        $out = & pwsh -NoProfile -File (Join-Path $script:Root 'Invoke-MailboxMessageReport.ps1') -ConfigPath $path 2>&1
        $LASTEXITCODE | Should -Be 1
        ($out -join ' ') | Should -Match 'Give the mailbox'
        Remove-Item $path
    }
}
