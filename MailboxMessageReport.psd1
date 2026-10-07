#
#  Mailbox Message Report - module manifest
#  --------------------------------------------------------------------------
#  Author  : Nicolas Fabert
#  Version : see ModuleVersion
#
#  Loaded by Invoke-MailboxMessageReport.ps1 (Import-Module by path).
#
@{
    RootModule        = 'MailboxMessageReport.psm1'
    ModuleVersion     = '1.1.0'
    GUID              = '9d3c5f41-7a2e-4b8c-a1d6-5e0f2c7b9a34'
    Author            = 'Nicolas Fabert'
    Copyright         = '(c) 2026 Nicolas Fabert. MIT License.'
    Description       = 'Mailbox Message Report: lists the messages of Exchange Online mailboxes - the primary mailbox and the In-Place Archive (read through Microsoft Graph as MBX:<ArchiveGuid>), Recoverable Items when asked - over a period or for a subject, with the date, subject, sender, recipients, Internet message ID and folder path of each message; CSV, JSON and HTML reports, global or per mailbox, and a window. Read only.'
    PowerShellVersion = '7.4'

    # Functions called by Invoke-MailboxMessageReport.ps1, the tests and the documentation tools. The other functions stay internal.
    FunctionsToExport = @(
        'Import-MmrConfiguration', 'Test-MmrConfiguration', 'New-MmrRequest', 'Test-MmrRequest', 'Read-MmrMailboxFile', 'Connect-MmrGraph'
        'Resolve-MmrMailboxes', 'Get-MmrFolders', 'Read-MmrMessages', 'Find-MmrMessages', 'New-MmrRunFolder', 'Export-MmrReport', 'Show-MmrGui', 'New-MmrForm'
        'Start-MmrLog', 'Stop-MmrLog', 'Write-MmrLog', 'Write-MmrBanner', 'Write-MmrStep', 'Write-MmrItem', 'Write-MmrSummary', 'Initialize-MmrSteps', 'Write-MmrNextStep'
        'Write-MmrRunBanner', 'Write-MmrMailboxTable', 'Write-MmrRunSummary', 'Format-MmrDuration'
    )
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()
    PrivateData       = @{
        PSData = @{
            Tags       = @('ExchangeOnline', 'Mailbox', 'Archive', 'InPlaceArchive', 'MicrosoftGraph', 'Report')
            LicenseUri = 'https://opensource.org/licenses/MIT'
        }
    }
}
