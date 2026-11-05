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
    Version : 2.1.1
    History : see CHANGELOG.md
#>
#Requires -Version 7.4
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Net.Http

$script:ToolVersion = '2.1.1'
$script:ToolRoot = $PSScriptRoot
# Package integrity. A package carries the owner-signed PackageIntegrity.psd1 manifest; when it is present, the signed
# verifier (src\Test-PackageIntegrity.ps1) must accept the tree before any part of the tool is compiled or loaded. A
# source tree without a manifest is an unverified development tree: the entry script refuses it unless -DevelopmentMode.
function Assert-MmrPackageIntegrity {
    [CmdletBinding()]
    param([string]$Root = $script:ToolRoot, [switch]$Distribution, [switch]$DevelopmentMode)
    $packageRoot = [IO.Path]::TrimEndingDirectorySeparator([IO.Path]::GetFullPath($Root))
    $manifest = Join-Path $packageRoot 'PackageIntegrity.psd1'
    if ((Test-Path -LiteralPath $manifest) -and -not (Test-Path -LiteralPath $manifest -PathType Leaf)) {
        throw 'Package integrity: PackageIntegrity.psd1 must be a file.'
    }
    if (-not (Test-Path -LiteralPath $manifest -PathType Leaf)) {
        # Without a manifest nothing is verified: say so once per process.
        if (-not [AppDomain]::CurrentDomain.GetData('MailboxMessageReport.UnverifiedTreeWarned')) {
            [AppDomain]::CurrentDomain.SetData('MailboxMessageReport.UnverifiedTreeWarned', $true)
            Write-Warning 'Package integrity: no PackageIntegrity.psd1; this is an unverified source tree (package byte integrity and publisher authentication are not established).'
        }
        return
    }
    $verifier = Join-Path $packageRoot 'src\Test-PackageIntegrity.ps1'
    $cursor = [IO.Path]::GetFullPath($verifier)
    while ($cursor) {
        if (-not (Test-Path -LiteralPath $cursor)) { throw 'Package integrity: the verifier is missing; nothing was loaded.' }
        $item = Microsoft.PowerShell.Management\Get-Item -LiteralPath $cursor -Force -ErrorAction Stop
        if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) { throw "Package integrity: reparse point: $cursor" }
        $cursor = [IO.Path]::GetDirectoryName($cursor)
    }
    $signature = Microsoft.PowerShell.Security\Get-AuthenticodeSignature -LiteralPath $verifier -ErrorAction Stop
    if ([string]$signature.Status -cne 'Valid' -or $null -eq $signature.SignerCertificate -or
        $signature.SignerCertificate.Thumbprint -cne '0F28E442889F17B5F4D251E8F96E7CC6E2C9C6D6' -or
        $null -eq $signature.TimeStamperCertificate) {
        throw 'Package integrity: verifier must have the trusted, timestamped owner signature. Use the complete release ZIP (https://github.com/Nico77600/MailboxMessageReport/releases/latest): download it again, extract it and unblock it.'
    }
    $arguments = @{ Root = $packageRoot }
    if ($Distribution) { $arguments.Distribution = $true }
    if ($DevelopmentMode) { $arguments.DevelopmentMode = $true }
    $result = @(& $verifier @arguments)
    if ($result.Count -ne 1 -or -not $result[0].PSObject.Properties['Status'] -or $result[0].Status -notlike 'Valid*') {
        throw 'Package integrity: the verifier did not return an explicit Valid result; the tool is not loaded.'
    }
}
Assert-MmrPackageIntegrity
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
