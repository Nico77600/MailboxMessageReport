<#
.SYNOPSIS
    Mailbox Message Report - configuration, request and dates (dot-sourced by MailboxMessageReport.psm1).

.DESCRIPTION
    The configuration file has sections (Tenant, Authentication, Search, Graph, Report, Window, Logging), like the other
    tools. It is flattened into one settings hashtable used by the command line, the window and the tests; unknown
    sections or keys and invalid values are all reported at once.

    A request is what one run reads: the mailboxes (typed, or a file that may give the ArchiveGuid of each one), the
    period, the subjects, where (primary mailbox, archive, Recoverable Items), the folders left out, and the report
    (formats, global or per mailbox). The command line and the window both build one with New-MmrRequest, so they are
    checked by the same rules (Test-MmrRequest).

    Dates: the period is typed in the time zone of Report.TimeZone (empty = the one of Windows). A start date is the
    start of that day; an end date without a time is the end of that day (included). No start and no end: every
    message, whatever its date.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.1.1
#>

# Section.Key of the configuration file -> key of the settings hashtable.
$script:ConfigSchema = [ordered]@{
    Tenant         = [ordered]@{ TenantId = 'TenantId'; Organization = 'Organization' }
    Authentication = [ordered]@{ Mode = 'AuthMode'; AppId = 'AppId'; CertificateThumbprint = 'CertificateThumbprint'; ClientSecretVariable = 'ClientSecretVariable' }
    Search         = [ordered]@{ Locations = 'Locations'; RecoverableItems = 'RecoverableItems'; Recipients = 'Recipients'; PastDays = 'PastDays'; ExcludeFolders = 'ExcludeFolders'; MailboxFile = 'MailboxFile' }
    Graph          = [ordered]@{ MaxConcurrency = 'MaxConcurrency'; PageSize = 'PageSize'; SplitFolderItems = 'SplitFolderItems'; MaxRetries = 'MaxRetries'; TimeoutSeconds = 'TimeoutSeconds' }
    Report         = [ordered]@{ OutputPath = 'OutputPath'; FilePrefix = 'ReportPrefix'; Formats = 'ReportFormats'; Layout = 'ReportLayout'; HtmlMaxMessages = 'HtmlMaxMessages'; CsvDelimiter = 'CsvDelimiter'; TimeZone = 'TimeZone' }
    Window         = [ordered]@{ PreviewMessages = 'PreviewMessages'; PreviewPerFolder = 'PreviewPerFolder'; ReadBody = 'ReadBody' }
    Logging        = [ordered]@{ Path = 'LogPath'; RetentionDays = 'LogRetentionDays' }
}

$script:Locations = @('Primary', 'Archive')
$script:Layouts = @('Global', 'PerMailbox', 'Both')
$script:GuidPattern = '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$'
$script:SmtpPattern = '^[^@\s<>"]+@[^@\s<>"]+\.[^@\s<>"]+$'
$script:EmptyGuid = '00000000-0000-0000-0000-000000000000'

function Get-MmrDefaultConfiguration {
    @{
        TenantId              = ''
        Organization          = ''
        AuthMode              = 'Certificate'
        AppId                 = ''
        CertificateThumbprint = ''
        ClientSecretVariable  = 'MMR_CLIENT_SECRET'
        Locations             = @('Primary', 'Archive')
        RecoverableItems      = $false
        Recipients            = $true
        PastDays              = 0
        ExcludeFolders        = @()
        MailboxFile           = ''
        MaxConcurrency        = 16
        PageSize              = 250
        SplitFolderItems      = 5000
        MaxRetries            = 6
        TimeoutSeconds        = 120
        OutputPath            = '.\reports'
        ReportPrefix          = 'MailboxMessageReport'
        ReportFormats         = @('Csv', 'Html')
        ReportLayout          = 'Global'
        HtmlMaxMessages       = 20000
        CsvDelimiter          = ';'
        TimeZone              = ''
        PreviewMessages       = 5000
        PreviewPerFolder      = 10
        ReadBody              = $true
        LogPath               = '.\logs'
        LogRetentionDays      = 30
    }
}

function Resolve-MmrPath {
    param([Parameter(Mandatory = $true)][string]$Path, [Parameter(Mandatory = $true)][string]$Root)
    if ([IO.Path]::IsPathRooted($Path)) { return [IO.Path]::GetFullPath($Path) }
    return [IO.Path]::GetFullPath($Path, $Root)
}

$script:ZoneCache = @{}

function Get-MmrTimeZone {
    <# Time zone of the dates: Report.TimeZone (IANA or Windows ID), the one of Windows when empty. Looked up once per ID. #>
    param([AllowEmptyString()][AllowNull()][string]$Id)
    if ([string]::IsNullOrWhiteSpace($Id)) { return [TimeZoneInfo]::Local }
    $zone = $script:ZoneCache[$Id]
    if (-not $zone) { $zone = [TimeZoneInfo]::FindSystemTimeZoneById($Id); $script:ZoneCache[$Id] = $zone }
    return $zone
}

function Format-MmrDate {
    <# A UTC date shown in the time zone of the report: yyyy-MM-dd HH:mm (or yyyy-MM-dd). -PeriodEnd: an end at 00:00 shows the day before (included). #>
    param([AllowNull()][object]$Utc, [AllowEmptyString()][AllowNull()][string]$TimeZone, [switch]$DateOnly, [switch]$PeriodEnd)
    return [MailboxMessageReportNative.Fast]::FormatDate($Utc, (Get-MmrTimeZone $TimeZone), [bool]$DateOnly, [bool]$PeriodEnd, $false)
}

function ConvertTo-MmrUtc {
    <#
        A date typed by the administrator (time zone of the report) to UTC. -EndOfDay: a date without a time is the
        end of that day (the next day at 00:00), so that the end date is included.
    #>
    param([Parameter(Mandatory = $true)][datetime]$Date, [AllowEmptyString()][AllowNull()][string]$TimeZone, [switch]$EndOfDay)
    if ($Date.Kind -eq [DateTimeKind]::Utc) { return $Date }
    $d = [datetime]::SpecifyKind($Date, [DateTimeKind]::Unspecified)
    if ($EndOfDay -and $d.TimeOfDay -eq [TimeSpan]::Zero) { $d = $d.AddDays(1) }
    return [TimeZoneInfo]::ConvertTimeToUtc($d, (Get-MmrTimeZone $TimeZone))
}

function Test-MmrConfiguration {
    <# Checks a settings hashtable (flattened configuration) and lists every problem. -ForConnection: tenant and application required. #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][hashtable]$Configuration, [switch]$ForConnection)

    $problems = [Collections.Generic.List[string]]::new()
    $c = $Configuration
    $number = {
        param([string]$Key, [string]$Label, [int]$Min, [int]$Max)
        $n = 0
        if (-not [int]::TryParse([string]$c[$Key], [ref]$n) -or $n -lt $Min -or $n -gt $Max) { [void]$problems.Add("$Label must be a whole number between $Min and $Max.") }
    }

    if ([string]$c.TenantId -and [string]$c.TenantId -notmatch $script:GuidPattern -and [string]$c.TenantId -notmatch '^[A-Za-z0-9-]+(\.[A-Za-z0-9-]+)+$') {
        [void]$problems.Add('Tenant.TenantId must be the tenant ID (GUID) or a domain of the tenant (contoso.onmicrosoft.com).')
    }
    if ([string]$c.AuthMode -notin 'Certificate', 'ClientSecret') { [void]$problems.Add("Authentication.Mode must be 'Certificate' or 'ClientSecret'.") }
    if ([string]$c.AppId -and [string]$c.AppId -notmatch $script:GuidPattern) { [void]$problems.Add('Authentication.AppId must be the application (client) ID, a GUID.') }
    if ([string]$c.AuthMode -eq 'Certificate' -and [string]$c.CertificateThumbprint -and [string]$c.CertificateThumbprint -notmatch '^[0-9a-fA-F]{40}$') {
        [void]$problems.Add('Authentication.CertificateThumbprint must be the 40 hexadecimal characters of the thumbprint.')
    }
    if ([string]$c.ClientSecretVariable -notmatch '^[A-Za-z_][A-Za-z0-9_]*$') { [void]$problems.Add('Authentication.ClientSecretVariable must be the name of an environment variable.') }
    if ($ForConnection) {
        if (-not [string]$c.TenantId) { [void]$problems.Add('Tenant.TenantId is required.') }
        if (-not [string]$c.AppId) { [void]$problems.Add('Authentication.AppId is required.') }
        if ([string]$c.AuthMode -eq 'Certificate' -and -not [string]$c.CertificateThumbprint) { [void]$problems.Add('Authentication.CertificateThumbprint is required in Certificate mode.') }
    }
    $locations = @($c.Locations)
    if ($locations.Count -eq 0 -or @($locations | Where-Object { $_ -notin $script:Locations }).Count) { [void]$problems.Add("Search.Locations must contain 'Primary', 'Archive' or both.") }
    if ($c.RecoverableItems -isnot [bool]) { [void]$problems.Add('Search.RecoverableItems must be $true or $false.') }
    if ($c.Recipients -isnot [bool]) { [void]$problems.Add('Search.Recipients must be $true or $false.') }
    & $number 'PastDays' 'Search.PastDays' 0 36500
    foreach ($p in @($c.ExcludeFolders)) { if ([string]$p -notmatch '^\\') { [void]$problems.Add("Search.ExcludeFolders: '$p' must be a folder path starting with \ (\Junk Email, \Inbox\Newsletters*).") } }
    & $number 'MaxConcurrency' 'Graph.MaxConcurrency' 1 32
    & $number 'PageSize' 'Graph.PageSize' 10 1000
    & $number 'SplitFolderItems' 'Graph.SplitFolderItems' 0 10000000
    & $number 'MaxRetries' 'Graph.MaxRetries' 0 10
    & $number 'TimeoutSeconds' 'Graph.TimeoutSeconds' 10 600
    foreach ($pair in @(@('OutputPath', 'Report.OutputPath'), @('ReportPrefix', 'Report.FilePrefix'), @('LogPath', 'Logging.Path'))) {
        if ([string]::IsNullOrWhiteSpace([string]$c[$pair[0]])) { [void]$problems.Add("$($pair[1]) is required.") }
    }
    if ([string]$c.ReportPrefix -match '[\\/:*?"<>|\s]') { [void]$problems.Add('Report.FilePrefix must be a file name without spaces.') }
    $formats = @($c.ReportFormats)
    if ($formats.Count -eq 0 -or @($formats | Where-Object { $_ -notin 'Csv', 'Html' }).Count) { [void]$problems.Add("Report.Formats must contain 'Csv', 'Html' or both.") }
    if ([string]$c.ReportLayout -notin $script:Layouts) { [void]$problems.Add("Report.Layout must be one of: $($script:Layouts -join ', ').") }
    & $number 'HtmlMaxMessages' 'Report.HtmlMaxMessages' 0 200000
    if ([string]$c.CsvDelimiter -notin ';', ',', "`t") { [void]$problems.Add("Report.CsvDelimiter must be ';', ',' or a tab.") }
    if ([string]$c.TimeZone) { try { $null = Get-MmrTimeZone $c.TimeZone } catch { [void]$problems.Add("Report.TimeZone '$($c.TimeZone)' is not a time zone of this computer (for example Europe/Paris, or empty).") } }
    & $number 'PreviewMessages' 'Window.PreviewMessages' 0 50000
    & $number 'PreviewPerFolder' 'Window.PreviewPerFolder' 0 1000
    if ($c.ReadBody -isnot [bool]) { [void]$problems.Add('Window.ReadBody must be $true or $false.') }
    & $number 'LogRetentionDays' 'Logging.RetentionDays' 1 365

    [pscustomobject]@{ IsValid = $problems.Count -eq 0; Problems = @($problems) }
}

function Import-MmrConfiguration {
    <#
        Reads config\MailboxMessageReport.config.psd1 (sections), applies the defaults, resolves the relative paths from
        the tool folder, checks everything and returns the settings hashtable.
    #>
    [CmdletBinding()]
    param(
        [string]$Path = (Join-Path $script:ToolRoot 'config\MailboxMessageReport.config.psd1'),
        [string]$Root = $script:ToolRoot
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "Configuration file not found: $Path" }
    $settings = Get-MmrDefaultConfiguration
    $problems = [Collections.Generic.List[string]]::new()
    $data = Import-PowerShellDataFile -LiteralPath $Path
    foreach ($section in $data.Keys) {
        if (-not $script:ConfigSchema.Contains($section)) { [void]$problems.Add("Unknown section '$section'. Sections: $($script:ConfigSchema.Keys -join ', ')."); continue }
        if ($data[$section] -isnot [hashtable]) { [void]$problems.Add("Section '$section' must be a @{ } block."); continue }
        foreach ($key in $data[$section].Keys) {
            if (-not $script:ConfigSchema[$section].Contains($key)) {
                [void]$problems.Add("Unknown key '$section.$key'. Keys of $($section): $($script:ConfigSchema[$section].Keys -join ', ').")
                continue
            }
            $settings[$script:ConfigSchema[$section][$key]] = $data[$section][$key]
        }
    }
    foreach ($key in 'Locations', 'ExcludeFolders', 'ReportFormats') { $settings[$key] = @($settings[$key] | Where-Object { "$_" }) }
    foreach ($key in 'OutputPath', 'LogPath', 'MailboxFile') {
        if (-not [string]::IsNullOrWhiteSpace([string]$settings[$key])) { $settings[$key] = Resolve-MmrPath -Path ([string]$settings[$key]) -Root $Root }
    }
    $settings.ConfigPath = [IO.Path]::GetFullPath($Path)
    foreach ($p in (Test-MmrConfiguration -Configuration $settings).Problems) { [void]$problems.Add($p) }
    if ($problems.Count) { throw ("Invalid configuration ($Path):`n - " + ($problems -join "`n - ")) }
    return $settings
}

function Read-MmrMailboxFile {
    <#
        The mailboxes of a file: a text file with one address per line (# = comment), or a CSV file (comma or semicolon)
        with a column PrimarySmtpAddress, EmailAddress, Mail, WindowsEmailAddress, UserPrincipalName, Address or
        Mailbox, and optionally ArchiveGuid (Get-EXOMailbox -Properties ArchiveGuid | Export-Csv): the archive is then
        read without User.Read.All. Returns @{ Address; ArchiveGuid } per mailbox, each address once.
    #>
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "Mailbox file not found: $Path" }
    $lines = @([IO.File]::ReadAllLines($Path) | ForEach-Object { $_.Trim() } | Where-Object { $_ -and -not $_.StartsWith('#') })
    if (-not $lines.Count) { return @() }
    $columns = @('PrimarySmtpAddress', 'EmailAddress', 'Mail', 'WindowsEmailAddress', 'UserPrincipalName', 'Address', 'Mailbox')
    $header = @($lines[0] -split '[;,]' | ForEach-Object { $_.Trim().Trim('"') })
    $found = @($columns | Where-Object { $header -contains $_ })
    $list = [Collections.Generic.List[object]]::new()
    $seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    if ($found.Count) {
        $delimiter = if ($lines[0].Contains(';')) { ';' } else { ',' }
        $hasArchive = $header -contains 'ArchiveGuid'
        foreach ($row in @($lines | ConvertFrom-Csv -Delimiter $delimiter)) {
            # The first address column of the row that holds a value.
            $address = $found | ForEach-Object { ([string]$row.$_).Trim() } | Where-Object { $_ } | Select-Object -First 1
            if (-not $address -or -not $seen.Add($address)) { continue }
            $guid = if ($hasArchive) { ([string]$row.ArchiveGuid).Trim().Trim('{', '}') } else { '' }
            if ($guid -eq $script:EmptyGuid) { $guid = '' }
            $list.Add([pscustomobject]@{ Address = $address.ToLowerInvariant(); ArchiveGuid = $guid.ToLowerInvariant() })
        }
    }
    else {
        foreach ($l in $lines) {
            $address = ($l -split '[;,\s]')[0].Trim('"')
            if ($address -and $seen.Add($address)) { $list.Add([pscustomobject]@{ Address = $address.ToLowerInvariant(); ArchiveGuid = '' }) }
        }
    }
    return $list.ToArray()
}

function Split-MmrList {
    <# Values typed in one text: separated by new lines, or by ; when -Semicolon. Trimmed, empty ones removed, each once. #>
    param([AllowEmptyString()][AllowNull()][string[]]$Text, [switch]$Semicolon, [switch]$Spaces)
    $pattern = if ($Spaces) { '[;,\s]+' } elseif ($Semicolon) { '[;\r\n]+' } else { '[\r\n]+' }
    $all = foreach ($t in @($Text)) { foreach ($v in ([string]$t -split $pattern)) { $v = $v.Trim(); if ($v) { $v } } }
    return @($all | Select-Object -Unique)
}

function New-MmrRequest {
    <#
        What one run reads, from the command line or the window, with the defaults of the configuration.
        Mailboxes: the addresses typed (-Mailbox) and those of the file (-MailboxFile, or Search.MailboxFile), each once;
        a file may give the ArchiveGuid of a mailbox.
    #>
    param(
        [Parameter(Mandatory = $true)][hashtable]$Settings,
        [string[]]$Mailbox,
        [string]$MailboxFile,
        [Nullable[datetime]]$Start,
        [Nullable[datetime]]$End,
        [string[]]$Subject,
        [string[]]$Location,
        [Nullable[bool]]$RecoverableItems,
        [Nullable[bool]]$Recipients,
        [string[]]$ExcludeFolder,
        [string]$Layout,
        [string[]]$Formats
    )

    $file = if ($MailboxFile) { [IO.Path]::GetFullPath($MailboxFile, (Get-Location).Path) } elseif (-not @($Mailbox | Where-Object { $_ }).Count) { [string]$Settings.MailboxFile } else { '' }
    $entries = [Collections.Generic.List[object]]::new()
    $seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($a in (Split-MmrList $Mailbox -Spaces)) { if ($seen.Add($a)) { $entries.Add([pscustomobject]@{ Address = $a.ToLowerInvariant(); ArchiveGuid = '' }) } }
    if ($file -and (Test-Path -LiteralPath $file -PathType Leaf)) {
        foreach ($e in (Read-MmrMailboxFile -Path $file)) { if ($seen.Add($e.Address)) { $entries.Add($e) } }
    }
    $startUtc = $null
    if ($null -ne $Start) { $startUtc = ConvertTo-MmrUtc -Date ([datetime]$Start) -TimeZone $Settings.TimeZone }
    elseif ([int]$Settings.PastDays -gt 0) {
        $today = [TimeZoneInfo]::ConvertTimeFromUtc([datetime]::UtcNow, (Get-MmrTimeZone $Settings.TimeZone)).Date
        $startUtc = ConvertTo-MmrUtc -Date $today.AddDays(-[int]$Settings.PastDays) -TimeZone $Settings.TimeZone
    }
    $endUtc = if ($null -ne $End) { ConvertTo-MmrUtc -Date ([datetime]$End) -TimeZone $Settings.TimeZone -EndOfDay } else { $null }
    [pscustomobject]@{
        Mailboxes        = $entries.ToArray()
        MailboxFile      = $file
        Start            = $startUtc
        End              = $endUtc
        Subject          = @(Split-MmrList $Subject)
        Location         = if ($Location) { @($script:Locations | Where-Object { $Location -contains $_ }) + @($Location | Where-Object { $_ -notin $script:Locations } | Select-Object -Unique) } else { @($Settings.Locations) }
        RecoverableItems = if ($null -ne $RecoverableItems) { [bool]$RecoverableItems } else { [bool]$Settings.RecoverableItems }
        Recipients       = if ($null -ne $Recipients) { [bool]$Recipients } else { [bool]$Settings.Recipients }
        ExcludeFolders   = if ($PSBoundParameters.ContainsKey('ExcludeFolder')) { @($ExcludeFolder | Where-Object { $_ }) } else { @($Settings.ExcludeFolders) }
        Layout           = if ($Layout) { $Layout } else { [string]$Settings.ReportLayout }
        Formats          = if ($Formats) { @($Formats | Select-Object -Unique) } else { @($Settings.ReportFormats) }
        TimeZone         = [string]$Settings.TimeZone
    }
}

function Test-MmrRequest {
    <# Checks a request (mailboxes, period, subjects, locations, report) and lists every problem. #>
    param([Parameter(Mandatory = $true)][pscustomobject]$Request)

    $problems = [Collections.Generic.List[string]]::new()
    $r = $Request
    if ($r.MailboxFile -and -not (Test-Path -LiteralPath $r.MailboxFile -PathType Leaf)) { [void]$problems.Add("Mailbox file not found: $($r.MailboxFile)") }
    elseif (-not @($r.Mailboxes).Count) {
        if ($r.MailboxFile) { [void]$problems.Add("No address in $($r.MailboxFile): one per line, or a CSV column PrimarySmtpAddress, UserPrincipalName, Mail or Address.") }
        else { [void]$problems.Add('Give the mailbox (-Mailbox) or a list of mailboxes (-MailboxFile): SMTP address or UPN.') }
    }
    foreach ($m in @($r.Mailboxes)) {
        if ($m.Address -notmatch $script:SmtpPattern) { [void]$problems.Add("Mailbox '$($m.Address)' is not an SMTP address or a UPN.") }
        if ($m.ArchiveGuid -and $m.ArchiveGuid -notmatch $script:GuidPattern) { [void]$problems.Add("Mailbox '$($m.Address)': ArchiveGuid '$($m.ArchiveGuid)' is not a GUID.") }
    }
    if ($null -ne $r.Start -and $null -ne $r.End -and $r.End -le $r.Start) { [void]$problems.Add('The end of the period must be after its start.') }
    $locations = @($r.Location)
    if (-not $locations.Count -or @($locations | Where-Object { $_ -notin $script:Locations }).Count) { [void]$problems.Add("Location: 'Primary', 'Archive' or both.") }
    foreach ($s in @($r.Subject)) {
        if ($s.Length -gt 255) { [void]$problems.Add("Subject '$($s.Substring(0, 40))...' is longer than 255 characters.") }
        if ($s -match '[\x00-\x1F]') { [void]$problems.Add('A subject cannot hold control characters.') }
    }
    foreach ($p in @($r.ExcludeFolders)) { if ([string]$p -notmatch '^\\') { [void]$problems.Add("Folder left out '$p' must be a path starting with \ (\Junk Email, \Inbox\Newsletters*).") } }
    if ([string]$r.Layout -notin $script:Layouts) { [void]$problems.Add("Report layout must be one of: $($script:Layouts -join ', ').") }
    $formats = @($r.Formats)
    if (-not $formats.Count -or @($formats | Where-Object { $_ -notin 'Csv', 'Html' }).Count) { [void]$problems.Add("Report formats: 'Csv', 'Html' or both.") }
    [pscustomobject]@{ IsValid = $problems.Count -eq 0; Problems = @($problems) }
}
