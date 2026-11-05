<#
.SYNOPSIS
    Mailbox Message Report - lists the messages of Exchange Online mailboxes: the primary mailbox and the In-Place
    Archive (and Recoverable Items when asked), through Microsoft Graph only.

.DESCRIPTION
    For each mailbox given (an address, a list, or a CSV file):
      - the user and its archive: Microsoft Graph gives the ID of the archive (MBX:<ArchiveGuid>@<tenant ID>), and
        the archive is then read like a mailbox of its own - no Exchange Online PowerShell. A CSV file may give the
        ArchiveGuid of each mailbox instead (Get-EXOMailbox -Properties ArchiveGuid | Export-Csv);
      - every mail folder of the primary mailbox and of the archive, with its path (\Inbox\Projects\Alpha), and the
        folders of Recoverable Items with -IncludeRecoverableItems;
      - the messages of each folder, over a period (received date) and for some subjects: date and time received and
        sent, subject, sender, recipients (To, Cc, Bcc), Internet message ID, folder and path, and whether the message
        is in the primary mailbox, the archive or Recoverable Items.

    Read only: nothing is ever changed in a mailbox, the body of a message is never read. Microsoft Graph,
    application permissions Mail.ReadBasic.All and User.Read.All (user guide, chapter 1). Writes CSV, JSON and HTML
    report files in a new folder - one report for every mailbox, one per mailbox, or both - and a daily log file.
    Everything is set in config\MailboxMessageReport.config.psd1; the parameters below override it.

.PARAMETER Mailbox
    The mailboxes: SMTP address (any alias) or UPN. Several may be given.

.PARAMETER MailboxFile
    A list of mailboxes: one address per line (# = comment), or a CSV file with a column PrimarySmtpAddress,
    UserPrincipalName, Mail, EmailAddress or Address, and optionally ArchiveGuid.

.PARAMETER Start
    The messages received from this date (date, or date and time), in the time zone of Report.TimeZone. Default: no
    limit (or today minus Search.PastDays).

.PARAMETER End
    The messages received until this date; a date without a time is included. Default: no limit.

.PARAMETER Subject
    Only the messages whose subject contains this text (any case). Several subjects: any of them.

.PARAMETER Location
    Primary, Archive or both (default: Search.Locations).

.PARAMETER IncludeRecoverableItems
    Also the folders of Recoverable Items (Deletions, Purges, Versions, DiscoveryHolds...), of the primary mailbox and
    of the archive. Default: Search.RecoverableItems.

.PARAMETER SkipRecipients
    Without To, Cc and Bcc: much faster on large folders of meeting messages (Exchange reads the recipients of each
    message for them). Default: Search.Recipients.

.PARAMETER ExcludeFolder
    Folders left out, by path with wildcards: '\Junk Email', '\Inbox\Newsletters*'. Default: Search.ExcludeFolders.

.PARAMETER Layout
    Global (one report for every mailbox), PerMailbox (one report per mailbox) or Both. Default: Report.Layout.

.PARAMETER Format
    Csv, Html or both. Default: Report.Formats. A Summary.json file is always written.

.PARAMETER HtmlMaxMessages
    Lowers the limit of messages of a HTML report for this run, 0 to 200,000 (the CSV file holds every message).
    To raise the limit, change Report.HtmlMaxMessages (500,000 delivered, up to 2,000,000).

.PARAMETER Gui
    Opens the window: the same search, the same report, a preview of the first messages.

.PARAMETER TenantId
    Overrides Tenant.TenantId.

.PARAMETER AppId
    Overrides Authentication.AppId.

.PARAMETER CertificateThumbprint
    Overrides Authentication.CertificateThumbprint.

.PARAMETER OutputPath
    Overrides Report.OutputPath.

.PARAMETER ConfigPath
    Configuration file. Default: config\MailboxMessageReport.config.psd1.

.PARAMETER DevelopmentMode
    Allows a source checkout that has no PackageIntegrity.psd1 manifest (a developer or continuous-integration
    build). A present manifest is always verified. The release ZIP is started without this switch.

.EXAMPLE
    .\Invoke-MailboxMessageReport.ps1 -Mailbox megan.bowen@contoso.com
    Every message of Megan Bowen, in her primary mailbox and her archive.

.EXAMPLE
    .\Invoke-MailboxMessageReport.ps1 -Mailbox megan.bowen@contoso.com -Location Archive -Start 2019-01-01 -End 2019-12-31
    The messages of her archive received in 2019.

.EXAMPLE
    .\Invoke-MailboxMessageReport.ps1 -MailboxFile C:\Tools\MailboxMessageReport-data\team.csv -Subject 'Contract Alpha', 'Project Alpha' -Layout Both
    The messages of every mailbox of the list whose subject contains one of the two texts: one report for all, and one
    per mailbox.

.EXAMPLE
    .\Invoke-MailboxMessageReport.ps1 -Mailbox john.doe@contoso.com -Subject 'Invoice' -IncludeRecoverableItems
    Also the messages deleted from Deleted Items, or kept by a hold.

.EXAMPLE
    .\Invoke-MailboxMessageReport.ps1 -Gui

.NOTES
    Author  : Nicolas Fabert
    Version : 2.1.1
    Exit codes : 0 = completed, 1 = failed, 2 = finished with warnings (a mailbox or a folder not read...).
    Documentation : docs\MailboxMessageReport-UserGuide.html (user guide: prerequisites, everyday commands) and
                    docs\MailboxMessageReport-Guide.html (developer guide)

    Disclaimer:
    This Script is a Personal project.
    It's provided "AS-IS". It's not an official Microsoft product so no support can be expected from Microsoft.

    As any scripts you must read carefully the documentation and test it first in a Test environment before any run in Production.

    The archive is read through an ID that Microsoft Graph gives (a beta endpoint) and opens, but that its
    documentation does not describe for the archive: test it in your environment before production use.
    The reports contain personal data (addresses, subjects): store and share them accordingly.
#>

#Requires -Version 7.4
[CmdletBinding()]
param(
    [string[]]$Mailbox,
    [string]$MailboxFile,
    [datetime]$Start,
    [datetime]$End,
    [string[]]$Subject,
    [ValidateSet('Primary', 'Archive')]
    [string[]]$Location,
    [switch]$IncludeRecoverableItems,
    [switch]$SkipRecipients,
    [string[]]$ExcludeFolder,
    [ValidateSet('Global', 'PerMailbox', 'Both')]
    [string]$Layout,
    [ValidateSet('Csv', 'Html')]
    [string[]]$Format,
    [ValidateRange(0, 200000)]
    [int]$HtmlMaxMessages,
    [switch]$Gui,
    [string]$TenantId,
    [string]$AppId,
    [string]$CertificateThumbprint,
    [string]$OutputPath,
    [string]$ConfigPath = (Join-Path $PSScriptRoot 'config\MailboxMessageReport.config.psd1'),
    [switch]$DevelopmentMode
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
$previousCulture = [Threading.Thread]::CurrentThread.CurrentCulture
[Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]::GetCultureInfo('en-US')
$exitCode = 1
$moduleLoaded = $false

try {
    # The package is verified before any of its code runs: the verifier must carry the owner signature, then it
    # checks the signed manifest. A tree without a manifest is accepted only with -DevelopmentMode.
    $integrityHint = 'Use the complete release ZIP (https://github.com/Nico77600/MailboxMessageReport/releases/latest): download it again, extract it and unblock it. A development or CI build is started from its folder with pwsh -NoProfile -File Invoke-MailboxMessageReport.ps1 <parameters> -DevelopmentMode.'
    $integrityManifest = Join-Path $PSScriptRoot 'PackageIntegrity.psd1'
    if (Microsoft.PowerShell.Management\Test-Path -LiteralPath $integrityManifest) {
        $integrityVerifier = Join-Path $PSScriptRoot 'src\Test-PackageIntegrity.ps1'
        $integrityCursor = [IO.Path]::GetFullPath($integrityVerifier)
        while ($integrityCursor) {
            $integrityItem = Microsoft.PowerShell.Management\Get-Item -LiteralPath $integrityCursor -Force -ErrorAction Stop
            if ($integrityItem.Attributes -band [IO.FileAttributes]::ReparsePoint) {
                throw "Package integrity: reparse point: $integrityCursor"
            }
            $integrityCursor = [IO.Path]::GetDirectoryName($integrityCursor)
        }
        $integritySignature = Microsoft.PowerShell.Security\Get-AuthenticodeSignature -LiteralPath $integrityVerifier -ErrorAction Stop
        if ([string]$integritySignature.Status -cne 'Valid' -or $null -eq $integritySignature.SignerCertificate -or
            $integritySignature.SignerCertificate.Thumbprint -cne '0F28E442889F17B5F4D251E8F96E7CC6E2C9C6D6' -or
            $null -eq $integritySignature.TimeStamperCertificate) {
            throw "Package integrity: verifier must have the trusted, timestamped owner signature. $integrityHint"
        }
        if ($DevelopmentMode) { & $integrityVerifier -Root $PSScriptRoot -DevelopmentMode | Out-Null } else { & $integrityVerifier -Root $PSScriptRoot -Distribution | Out-Null }
    }
    elseif (-not $DevelopmentMode) {
        throw "Package integrity: PackageIntegrity.psd1 is missing. $integrityHint"
    }
    else {
        Write-Warning 'DevelopmentMode: no package manifest; package byte integrity and publisher authentication are not established.'
        # The module does not repeat this warning in the same process.
        [AppDomain]::CurrentDomain.SetData('MailboxMessageReport.UnverifiedTreeWarned', $true)
    }
    Import-Module (Join-Path $PSScriptRoot 'MailboxMessageReport.psd1') -Force
    $moduleLoaded = $true

    # ---- configuration, then command-line overrides ---------------------------------------------------
    $settings = Import-MmrConfiguration -Path $ConfigPath -Root $PSScriptRoot
    if ($TenantId) { $settings.TenantId = $TenantId }
    if ($AppId) { $settings.AppId = $AppId }
    if ($CertificateThumbprint) { $settings.CertificateThumbprint = $CertificateThumbprint; $settings.AuthMode = 'Certificate' }
    if ($OutputPath) { $settings.OutputPath = [IO.Path]::GetFullPath($OutputPath, (Get-Location).Path) }
    if ($PSBoundParameters.ContainsKey('HtmlMaxMessages')) { $settings.HtmlMaxMessages = $HtmlMaxMessages }
    $check = Test-MmrConfiguration -Configuration $settings
    if (-not $check.IsValid) { throw ("Invalid value:`n - " + ($check.Problems -join "`n - ")) }

    $logPath = Start-MmrLog -Directory $settings.LogPath -RetentionDays $settings.LogRetentionDays

    if ($Gui) {
        Write-MmrLog 'STEP' '=== Mailbox Message Report - window opened ==='
        Show-MmrGui -Configuration $settings
        $exitCode = 0
    }
    else {
        $requestArgs = @{ Settings = $settings; Mailbox = $Mailbox; MailboxFile = $MailboxFile; Subject = $Subject; Location = $Location; Layout = $Layout; Formats = $Format }
        if ($PSBoundParameters.ContainsKey('Start')) { $requestArgs.Start = $Start }
        if ($PSBoundParameters.ContainsKey('End')) { $requestArgs.End = $End }
        if ($PSBoundParameters.ContainsKey('IncludeRecoverableItems')) { $requestArgs.RecoverableItems = [bool]$IncludeRecoverableItems }
        if ($PSBoundParameters.ContainsKey('SkipRecipients')) { $requestArgs.Recipients = -not $SkipRecipients }
        if ($PSBoundParameters.ContainsKey('ExcludeFolder')) { $requestArgs.ExcludeFolder = $ExcludeFolder }
        $request = New-MmrRequest @requestArgs
        $requestCheck = Test-MmrRequest -Request $request
        if (-not $requestCheck.IsValid) { throw ("Cannot run:`n - " + ($requestCheck.Problems -join "`n - ")) }

        Write-MmrRunBanner -Settings $settings -Request $request -LogPath $logPath
        Initialize-MmrSteps -Total 5

        # ---- Microsoft Graph --------------------------------------------------------------------------
        Write-MmrNextStep 'Microsoft Graph' 'Key'
        $connection = Connect-MmrGraph -Settings $settings
        $dot = [char]0x00B7
        Write-MmrItem Ok ("Application {0} {1} tenant {2}" -f $(if ($connection.AppName) { "$($connection.AppName) ($($settings.AppId))" } else { $settings.AppId }), $dot, $connection.TenantGuid) -Icon Key
        Write-MmrItem Info ("Permissions: {0}" -f (@($connection.Roles) -join ', ')) -Icon Shield
        if (@($connection.Roles | Where-Object { $_ -in 'Mail.ReadWrite' }).Count) { Write-MmrItem Warn 'Mail.ReadWrite: more than needed (the tool only reads). Mail.ReadBasic.All is enough.' }

        # ---- mailboxes, folders, messages (part files in the folder of the run), then the report ------------
        $runPath = New-MmrRunFolder -OutputPath $settings.OutputPath -Prefix $settings.ReportPrefix
        $partsPath = Join-Path $runPath '.parts'
        try {
            $result = Find-MmrMessages -Settings $settings -Request $request -PartsPath $partsPath
            if ($result.Mailboxes.Count -gt 1) { Write-MmrMailboxTable -Mailboxes @($result.Mailboxes) }

            Write-MmrNextStep 'Report' 'Report'
            $report = Export-MmrReport -Result $result -Directory $runPath -Prefix $settings.ReportPrefix -Formats $request.Formats -Layout $request.Layout -Delimiter $settings.CsvDelimiter -HtmlMaxMessages $settings.HtmlMaxMessages -PartsPath $partsPath
        }
        catch {
            if (Test-Path -LiteralPath $partsPath) { Remove-Item -LiteralPath $partsPath -Recurse -Force -ErrorAction SilentlyContinue }
            throw
        }
        foreach ($f in $report.Files.Values) { Write-MmrItem Ok $f -Icon File }
        if ($report.MailboxFiles) { Write-MmrItem Ok ('{0:N0} mailbox report(s) in {1}' -f $report.MailboxFiles, (Join-Path $runPath 'Mailboxes')) -Icon File }
        $reportText = if ($report.Files.Contains('Html')) { $report.Files.Html } else { $report.Directory }
        Write-MmrRunSummary -Result $result -ReportText $reportText -LogPath $logPath
        $exitCode = switch ($result.Status) { 'Completed' { 0 } 'Failed' { 1 } default { 2 } }
    }
}
catch {
    if ($moduleLoaded) {
        Write-MmrItem Fail $_.Exception.Message
        Write-MmrLog 'ERROR' ($_.ScriptStackTrace -replace '\r?\n', ' | ')
        Write-Host ''
    }
    else {
        Write-Host "Mailbox Message Report: $($_.Exception.Message)" -ForegroundColor Red
    }
    $exitCode = 1
}
finally {
    if ($moduleLoaded) { Stop-MmrLog }
    [Threading.Thread]::CurrentThread.CurrentCulture = $previousCulture
}
exit $exitCode
