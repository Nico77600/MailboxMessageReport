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
    Version : 2.1.1
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
        SplitFolderItems      = 2000
        MaxRetries            = 6
        TimeoutSeconds        = 120
        OutputPath            = '.\reports'
        ReportPrefix          = 'MailboxMessageReport'
        ReportFormats         = @('Csv', 'Html')
        ReportLayout          = 'Global'
        HtmlMaxMessages       = 500000
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
    & $number 'HtmlMaxMessages' 'Report.HtmlMaxMessages' 0 2000000
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

# SIG # Begin signature block
# MIImIQYJKoZIhvcNAQcCoIImEjCCJg4CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCV6MY4sAj8SKqI
# F4pfCdxi4Fj294Bt/ezlUkKbai1a+aCCDW8wggZyMIIEWqADAgECAghkM1HTxzif
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
# wZislcdCZuS6MYIYCDCCGAQCAQEwgYwweDELMAkGA1UEBhMCVVMxDjAMBgNVBAgM
# BVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMREwDwYDVQQKDAhTU0wgQ29ycDE0MDIG
# A1UEAwwrU1NMLmNvbSBDb2RlIFNpZ25pbmcgSW50ZXJtZWRpYXRlIENBIFJTQSBS
# MQIQBof4ljXihJpmCqqt2mhAWzANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3
# AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisG
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCD0wS3aqbDE
# 08ysfftO0S6dYZaRIGUGBjzniWlLH3GGBDANBgkqhkiG9w0BAQEFAASCAgBS4cvH
# AhoUYRIs2+xiFS/1Gd9GLFntpWMd0eQJR3QsKmos9enNntpnOPCK0Zpd6pOFnI33
# IoQN2npNJOnIEBP7ixAYG74R6XdJsU0rmflLzuaAqX1YXzl9Q86NfFj0oQiBYdM9
# AS5zNGQ6y+I2ORe8hNM7iMaqmIDthpUoFBU64anEd7nccXpTaXDCSos4O4VFaEjA
# TV8PH6sFN3ntjmhLaa9v8FpmRqI310hMH/1JAsbGRrTJITtvpGlbD92vDWVdpNHy
# tum7mw5vUSV5eB3HlDjUdydZV13hXhrfRK4mvYqCd0NyCdWdEf5+87pWWFCYe8ln
# h3BHPXcb3RGtuJ7ETFP/c+cQzJJ+YeDVgAXs2uC0DnkW2odleHKKL8Aow7VeKl9x
# 8X0yH0kekcE3GnE0v/bgNtdT/zRrE7JHA7oEXu98/0PPHs1n2f5PjRspSgxC3YxN
# kAcyqPlYNAVMUKqODDomRkSxhC7Mqp9+sEh2DEBD/e5CbvI+2kdF/otMtSXqPJDl
# Z+hRRPHAbLUGo6sk2834wNhwWhHqlblUuY+LaTtvqVDr2Gum1/dTn4glRyQLbD1U
# v1zCzS/x/ETmSUyy7arae7fsDsoolKBMR7iRR0ir0ZD1sAkxIabeQz6KCv7oOs4D
# /61HT6RtVJJHTDntU1C7UsIYohYkjYlJsxBmzKGCFMUwghTBBgorBgEEAYI3AwMB
# MYIUsTCCFK0GCSqGSIb3DQEHAqCCFJ4wghSaAgEDMQ0wCwYJYIZIAWUDBAIBMIGB
# BgsqhkiG9w0BCRABBKByBHAwbgIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEINlNR3gWL0aEn3fwYuObxHcUm2CU/sheBEIJn9BWht4RAggrBhbcmsj3
# pRgPMjAyNjExMDUxNzUxMjBaMAMCAQECCB5Ix+f7yVthoIIRpzCCBPMwggLboAMC
# AQICEFRD+De9Wu4YAYRyuFzajuMwDQYJKoZIhvcNAQELBQAwcjELMAkGA1UEBhMC
# VVMxDjAMBgNVBAgMBVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMREwDwYDVQQKDAhT
# U0wgQ29ycDEuMCwGA1UEAwwlU1NMLmNvbSBUaW1lc3RhbXBpbmcgSXNzdWluZyBS
# U0EgQ0EgMjAeFw0yNjA2MTIxNzIyNDZaFw0zNjA2MDkxNzIyNDVaMG4xCzAJBgNV
# BAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91c3RvbjERMA8GA1UE
# CgwIU1NMIENvcnAxKjAoBgNVBAMMIVNTTC5jb20gVGltZXN0YW1waW5nIFVuaXQg
# MjAyNiBFMTBZMBMGByqGSM49AgEGCCqGSM49AwEHA0IABOgli7hmEy5IkjjzCzbJ
# DgbYV7CIUUgOOn3jkML0i3A2jYer8VMwShMdSnRQURbOzh65x/ytFQOjGw0CBaVI
# ND2jggFSMIIBTjAfBgNVHSMEGDAWgBRwmmXBQ4Cj/OmxZIlu8TCcCSHJfDBNBggr
# BgEFBQcBAQRBMD8wPQYIKwYBBQUHMAKGMWh0dHA6Ly9jZXJ0LnNzbC5jb20vU1NM
# LmNvbS10aW1lU3RhbXBpbmctSS1SMi5jZXIwUQYDVR0gBEowSDA8BgwrBgEEAYKp
# MAEDBgEwLDAqBggrBgEFBQcCARYeaHR0cHM6Ly93d3cuc3NsLmNvbS9yZXBvc2l0
# b3J5MAgGBmeBDAEEAjAWBgNVHSUBAf8EDDAKBggrBgEFBQcDCDBCBgNVHR8EOzA5
# MDegNaAzhjFodHRwOi8vY3Jscy5zc2wuY29tL1NTTC5jb20tdGltZVN0YW1waW5n
# LUktUjIuY3JsMB0GA1UdDgQWBBQnrANWuihquhc7gz1Af3moJOzDPDAOBgNVHQ8B
# Af8EBAMCB4AwDQYJKoZIhvcNAQELBQADggIBAE92bs78A57CGU1wZ/oY5rgV4Mwd
# dn22MU/bmHumb4b6W2VinMg0IGnBE5D8PJOsnqp2AmMhWdOt5TDxKgIQA9pzBute
# GYez7R1vQ1qn7IIr8Wz2AXjfre1MYJGE93saW1jHTxKgbVlchzry3CrzMYCVnyYw
# DcFZMFQ35Mv4+3H9vPoopjfDynO0v7vPy3p0UZ1dltmH53EZ5M/Qn7OZ0Q82mYLo
# h666dhscIJzUEh+BgYxzpdz8y7J1WxqGSouGRI4F7qbOJhhjiu/hgdhwJUJ/jkve
# OHB6pZB9IBJbdlns6AjUaRC6YfamsbgZ40hKl1ln4BOshgcDtLsSjaqVlUx5gMK+
# krwT6H8bm1yLhJ/Yr0eAizrgMVEqyeNWQYP6qhtKCYxehpESm8au3MlkpBGWCYZi
# goNh4OX1x2oDsYsG5p+ST2db4Vgq25N7/3U8doL0yufnqwuR/f603JY1CO++sndq
# 34Y35Kb8gLG7V7OA+5658PIvcKK+fn3bUUiQVcasy5spBX7aC9FKdkPvrciglHa4
# m3XJ6weBykt7PrdwI3XGwNOgwi/Bon0IvXy9MkQUTZQOKOgT/kq++eTGgjY1EWGQ
# 7w8sSRwQU+3Ko2SMLIkfLPuvOtJ6BWwQbXYHtiMOwjwv9UN9SNu4VaJUuwdMmqN7
# /F9jLVst5rGZV03BMIIF3TCCA8WgAwIBAgIIeyyb0xaAMpkwDQYJKoZIhvcNAQEL
# BQAwfDELMAkGA1UEBhMCVVMxDjAMBgNVBAgMBVRleGFzMRAwDgYDVQQHDAdIb3Vz
# dG9uMRgwFgYDVQQKDA9TU0wgQ29ycG9yYXRpb24xMTAvBgNVBAMMKFNTTC5jb20g
# Um9vdCBDZXJ0aWZpY2F0aW9uIEF1dGhvcml0eSBSU0EwHhcNMTYwMjEyMTczOTM5
# WhcNNDEwMjEyMTczOTM5WjB8MQswCQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMx
# EDAOBgNVBAcMB0hvdXN0b24xGDAWBgNVBAoMD1NTTCBDb3Jwb3JhdGlvbjExMC8G
# A1UEAwwoU1NMLmNvbSBSb290IENlcnRpZmljYXRpb24gQXV0aG9yaXR5IFJTQTCC
# AiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAPkP3aMrfcvQKv7sZ4Wm5y4b
# unfh4/WvpOz6Sl2RxFdHaxh3a3by/ZPkPQ/CFp4LZsNWlJ4Xg4XOVu/yFv0AYvUi
# CVToZRdOQbngT0aXqhvIuG5iXmmxX9sqAn78bMrzQdjt0Oj8P2FI7bADFB0QDksZ
# 4LtO7IZl/zbzXmcCC52GVWH9ejjt/uIZALdvoVBidXQ8oPrIJZK0bnoix/geoeOy
# 3ZExqysdBP+lSgQ36YWkMyv94tZVNHwZpEpox7Ko07fKoZOI68GXvIz5HdkihCR0
# xwQ9aqkpk8zruFvh/l8lqjRYyMEjVJ0bmBHDOJx+PYZspQ9AhnwC9FwCTyjLrnGf
# DzrIM/4RJTXq/LrFYD3ZfBjVsqnTdXgDciLKOsMf7yzlLqn6niy2UUb9rwPW6mBo
# 6oUWNmuF6R7As93EJNyAKoFBbZQ+yODJgUEAnl6/f8UImKIYLEJAs/lvOCdLToD0
# PYFH4Ih86hzOtXVcUS4cK38acijnALXRdMbX5J+tB5O2UzU1/Dfkw/ZdFr4hc96S
# CvigY2q8lpJqPvi8ZVWb3vUNiSYE/CUapiVpy8JtynziWV+XrOvvLsi81xtZPCvM
# 8hnIk2snYxnP/Okm+Mpxm3+T/jRnhE6Z6/yzeAkzcLpmpnbtG3PrGqUNxCITIJRW
# Ck4sbE6x/c+cCbqiM+2HAgMBAAGjYzBhMB0GA1UdDgQWBBTdBAkHovV6fVJTEpKV
# 7jiAJQ2mWTAPBgNVHRMBAf8EBTADAQH/MB8GA1UdIwQYMBaAFN0ECQei9Xp9UlMS
# kpXuOIAlDaZZMA4GA1UdDwEB/wQEAwIBhjANBgkqhkiG9w0BAQsFAAOCAgEAIBgR
# lCn7Jp0cHh5wYfGVcpNxJK1ok1iOMq8bs3AD/CUrdIWQPXhq9LmLpZc7tRiRux6n
# +UBbkflVma8eEdBcHadm47GUBwwyOabqG7B52B2ccETjit3E+ZUfijhDPwGFpUen
# PUayvOUiaPd7nNgsPgohyC0zrL/FgZkxdMF1ccW+sfAjRfSda/wZY52jvATGGAsl
# u1OJD7OAUN5F7kR/q5R4ZJjT9ijdh9hwZXT7DrkT66cPYakylszeu+1jTBi7qUD3
# oFRuIIhxdRjqerQ0cuAjJ3dctpDqhiVAq+8zD8ufgr6iIPv2tS0a5sKFsXQP+8hl
# AqRSAUfdSSLBv9jra6x+3uxjMxW3IwiPxg+NQVrdjsW5j+VFP3jbutIbQLH+cU0/
# 4IGiul607BXgk90IH37hVZkLId6Tngr75qNJvTYw/ud3sqB1l7UtgYgXZSD32pAA
# n8lSzDLKNXz1PQ/YK9f1JmzJBjSWFupwWRoyeXkLtoh/D1JIPb9s2KJELtFOt3JY
# 04kTlf5Eq/jXixtunLwsoFvVagCvXzfh1foQC5ichucmj87w7G6KVwuA406ywKBj
# YZC6VWg3dGq2ktufoYYitmUnDuy2n0Jg5GfCtdpBC8TTi2EbvPofkSvXRAdeuims
# 2cXp71NIWuuA8ShYIc2wBlX7Jz9TkHCpBB5XJ7kwggbLMIIEs6ADAgECAhBSnbvC
# uImH9YOA4LDrogY2MA0GCSqGSIb3DQEBCwUAMHwxCzAJBgNVBAYTAlVTMQ4wDAYD
# VQQIDAVUZXhhczEQMA4GA1UEBwwHSG91c3RvbjEYMBYGA1UECgwPU1NMIENvcnBv
# cmF0aW9uMTEwLwYDVQQDDChTU0wuY29tIFJvb3QgQ2VydGlmaWNhdGlvbiBBdXRo
# b3JpdHkgUlNBMB4XDTI2MDYxMjE2NDIxNVoXDTQxMDIxMjE3MzkzOVowcjELMAkG
# A1UEBhMCVVMxDjAMBgNVBAgMBVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMREwDwYD
# VQQKDAhTU0wgQ29ycDEuMCwGA1UEAwwlU1NMLmNvbSBUaW1lc3RhbXBpbmcgSXNz
# dWluZyBSU0EgQ0EgMjCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAJ8v
# q+ZzTZC8MUabNgd5WlBjVPcu71oiedNd8n9Il8+uW0HcGvW44/Jf0pHl6hxqn0od
# haGnvvNBPJgxH/d6CMI7Mtgl9nF5isarRfByRmAeH+skYShA+dMJGzsrC9NiF3/v
# xGZ2uOI334YVc+DlP5OdsUPz5yIP7hYXNlpPknPwh6gK/pWFZD8+y8uUagfX9USh
# q/SQt+1tS4FXRwudHTSHbGt0Fz4ZbWGbm0d56ROAW34tI6FFuZ+TtGssHjeOUf2F
# S8QNGBpadTh1X1Z7Kl6vG4Z8NS3OJIyelzoF5TpcP2vM3ue0723bmZd3oXsE5rzt
# AJMv56jlg0vSDz8XcXRur6awW0rh5cTQDQQGx04hYJH2Zn9HWWJU/wJ2SKBcQKtB
# yCz6JX+DoFoBE++kolkpZtl2ESF8FY6XqhtKZb18ESEAc0ohCY+0HkBZwrx8eIED
# CDbzd/shv9a8YRG2oxr3tLUz95V+Dzzh1QBNECaeGczyNWGCZIVUa4lVcVUVOBbI
# OAT8D+iPmC5h29bWeES8DhTS4f4CWA4D5QoQqMTqIBcV4MD2Waw8hEWAdGnjnsVt
# DimnDeEUbvXKhdiirvwG+QC+3mAy1RlD6xBMmGyDAKVSqqXm3NP4w0J4gvt65Grk
# QI4FXPJtYqfwkIBtQ8OGfbpVUQGPVUemWtui7E21AgMBAAGjggFRMIIBTTASBgNV
# HRMBAf8ECDAGAQH/AgEAMB8GA1UdIwQYMBaAFN0ECQei9Xp9UlMSkpXuOIAlDaZZ
# ME0GCCsGAQUFBwEBBEEwPzA9BggrBgEFBQcwAoYxaHR0cDovL2NlcnQuc3NsLmNv
# bS9TU0xjb20tUm9vdENBLVJTQS00MDk2LVIxLmNlcjA/BgNVHSAEODA2MDQGBFUd
# IAAwLDAqBggrBgEFBQcCARYeaHR0cHM6Ly93d3cuc3NsLmNvbS9yZXBvc2l0b3J5
# MBMGA1UdJQQMMAoGCCsGAQUFBwMIMEIGA1UdHwQ7MDkwN6A1oDOGMWh0dHA6Ly9j
# cmxzLnNzbC5jb20vU1NMY29tLVJvb3RDQS1SU0EtNDA5Ni1SMS5jcmwwHQYDVR0O
# BBYEFHCaZcFDgKP86bFkiW7xMJwJIcl8MA4GA1UdDwEB/wQEAwIBhjANBgkqhkiG
# 9w0BAQsFAAOCAgEAZRRbzTHpgOogYgWKOeiCkI1oVrkGmCrqOKqX9CiTREQLdYM/
# h7ANki98QP1fhAPEkTzT/Ur19IrPuZg7KaWL6SsmBqmD9+mBVtJmxdRdGbV8bpgl
# FDv2VeZgTSg3+UOD2O0G9HlgHqVHNKzSa/zL5+WYl3yT8Q29sZnIPGkQ+udv56/j
# q5NlAT9BON4iXSOCxe1bNvUd0hkS2GTsTNeQyOEHm7CkZdCg5X+DpmBginO7pk8Y
# m5GbTdQtSloJ3/Y2nVPycJ8kmmBAYMev7/hVQukhjjc4mO03LyPTFfqmMR1D5Tpu
# gGkD9nUuG41l55vOpXQm1JMbTgArpkn0/e9O+yTYlWuz2dSAZHf9jxNQAvILO1Pv
# 2JNQTfV0TyIteZP7gNcRfaHnnuSho8ohipKiE+1vEJIFE+c3gopxSpHCmn2THETk
# lWKV7ZrUKhhkdoOLj2KidCtYZrOoCV6x7f9+KfDPT61fKDVntGZdRKwzLV53VIN6
# oSS+KQN2VJaO8OZTOZpfgO4o6/9xpNr+EEaA3hNkoZk13R4LaqPC/309dfZmHBE8
# nBNjWBkV7AZdzdK8w7B0a/D9066a6oUa19OIktD+3kPpQ8YEt8oRp9zs1hoQ7fRn
# 9sdantP6nYRQFTpF8czGc+FGresbAH819Pv5vBhCluTo8Px2PpC71cM0jzsxggJV
# MIICUQIBATCBhjByMQswCQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNV
# BAcMB0hvdXN0b24xETAPBgNVBAoMCFNTTCBDb3JwMS4wLAYDVQQDDCVTU0wuY29t
# IFRpbWVzdGFtcGluZyBJc3N1aW5nIFJTQSBDQSAyAhBUQ/g3vVruGAGEcrhc2o7j
# MAsGCWCGSAFlAwQCAaCCAWAwGgYJKoZIhvcNAQkDMQ0GCyqGSIb3DQEJEAEEMBwG
# CSqGSIb3DQEJBTEPFw0yNjExMDUxNzUxMjBaMCgGCSqGSIb3DQEJNDEbMBkwCwYJ
# YIZIAWUDBAIBoQoGCCqGSM49BAMCMC8GCSqGSIb3DQEJBDEiBCCRQ5bDxVpLxsc9
# K+slxBDkvLpdkoIe8aCFCv1PCCinNzCByAYLKoZIhvcNAQkQAi8xgbgwgbUwgbIw
# ga8EIFdP6heq+l58aiubCMmYsXNeRQs4c8pE5qJMoJMOY/OwMIGKMHakdDByMQsw
# CQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNVBAcMB0hvdXN0b24xETAP
# BgNVBAoMCFNTTCBDb3JwMS4wLAYDVQQDDCVTU0wuY29tIFRpbWVzdGFtcGluZyBJ
# c3N1aW5nIFJTQSBDQSAyAhBUQ/g3vVruGAGEcrhc2o7jMAoGCCqGSM49BAMCBEYw
# RAIgTOtUJvcD70KSQbxxADzmA/lzPwELRV+eRfqbK5E4hl8CIDX6dZEf9tGLcYo2
# 0/1vEnUbeF3lJsPtQ+7Lv9o68kBE
# SIG # End signature block
