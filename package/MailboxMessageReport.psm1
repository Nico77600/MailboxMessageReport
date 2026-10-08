<#
.SYNOPSIS
    Mailbox Message Report - PowerShell module.

.DESCRIPTION
    Loads the parts of the tool, in the order of an execution:

        src\MailboxMessageReport.Console.ps1    console output and log file (same rules as the other tools)
        src\MailboxMessageReport.Config.ps1     configuration file, settings, request, dates
        src\MailboxMessageReport.Graph.ps1      Microsoft Graph: token, requests, $batch scheduler, page reader
        src\MailboxMessageReport.Mailboxes.ps1  the mailboxes: user, primary mailbox, archive (MBX:<ArchiveGuid>)
        src\MailboxMessageReport.Search.ps1     folders and messages, written to part files
        src\MailboxMessageReport.Report.ps1     CSV, JSON and HTML report, global or per mailbox
        src\MailboxMessageReport.Gui.ps1        WPF window (Fluent theme of Windows 11)

    The access token and the client secret stay in memory: they are never written to the console, the log or the
    report. The tool only reads: it never changes a mailbox.

.NOTES
    Author  : Nicolas Fabert
    Version : 2.1.0
    History : see CHANGELOG.md
#>
#Requires -Version 7.4
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Net.Http

$script:ToolVersion = '2.1.0'
$script:ToolRoot = $PSScriptRoot
$script:LogWriter = $null
$script:LogPath = $null
$script:Quiet = $false
# GUI hooks, set only while the window runs a search: Queue (progress lines, read by the window), Sink (lines, inline
# runs), Cancel, Pump.
$script:Ui = $null
# Microsoft Graph connection of the current run (Connect-MmrGraph).
$script:Graph = $null

# Compiled helpers (src\MailboxMessageReport.Native.cs): once per PowerShell process.
$native = 'MailboxMessageReportNative.Fast' -as [type]
if (-not $native) { Add-Type -Path (Join-Path $PSScriptRoot 'src\MailboxMessageReport.Native.cs') }
elseif ($native::Version -ne $script:ToolVersion) { throw "Mailbox Message Report $($native::Version) is already loaded in this PowerShell session: open a new PowerShell window to use $($script:ToolVersion)." }

foreach ($part in 'Console', 'Config', 'Graph', 'Mailboxes', 'Search', 'Report', 'Gui') {
    . (Join-Path $PSScriptRoot "src\MailboxMessageReport.$part.ps1")
}

Export-ModuleMember -Function @(
    'Import-MmrConfiguration', 'Test-MmrConfiguration', 'New-MmrRequest', 'Test-MmrRequest', 'Read-MmrMailboxFile', 'Connect-MmrGraph'
    'Resolve-MmrMailboxes', 'Get-MmrFolders', 'Read-MmrMessages', 'Find-MmrMessages', 'New-MmrRunFolder', 'Export-MmrReport', 'Show-MmrGui', 'New-MmrForm'
    'Start-MmrLog', 'Stop-MmrLog', 'Write-MmrLog', 'Write-MmrBanner', 'Write-MmrStep', 'Write-MmrItem', 'Write-MmrSummary', 'Initialize-MmrSteps', 'Write-MmrNextStep'
    'Write-MmrRunBanner', 'Write-MmrMailboxTable', 'Write-MmrRunSummary', 'Format-MmrDuration'
)
