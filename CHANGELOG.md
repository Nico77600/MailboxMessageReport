# Changelog

All notable changes to Mailbox Message Report. Versions: MAJOR.MINOR.PATCH (developer guide, appendix D).

## 2.1.1 - 2026-11-05

Signed PowerShell files, a release ZIP with its own README and a rewritten Get started.

### Added
- The PowerShell files of the release ZIP carry the owner's timestamped Authenticode signature (Nicolas Fabert).
- Before it starts, the tool checks the signature and SHA256 of every file of the package and stops if one is changed, missing or added. Only `config\MailboxMessageReport.config.psd1`, `reports\` and `logs\` may change in the tool folder; keep mailbox lists and certificate exports outside it.
- New parameter `-DevelopmentMode` starts a source checkout that has no signed file list.
- The user guide explains the Microsoft Entra registration step by step: permissions, admin consent and certificate, each with what you should see.
- The README and the script help state the project disclaimer.

### Changed
- The release ZIP has one folder, `MailboxMessageReport-2.1.1`, with the README, LICENSE, CHANGELOG, THIRD-PARTY-NOTICES, `PackageIntegrity.psd1` (the signed file list) and the two HTML guides.
- The README in the ZIP links only to files that are in the ZIP.
- Get started begins with the release ZIP and the PowerShell 7.4 check; the example folder is `C:\Tools\MailboxMessageReport-2.1.1`.
- Technical details (Graph calls, measurements) are in the developer guide only.
- The source archives keep the run-time files, the README, the CHANGELOG and the license.

## 2.1.0 - 2026-10-07

Large folders are read faster, and the progress bar shows the messages read.

- Large folders are read in several parts at the same time, each part holding the same number of messages.
- A folder of more than 2,000 messages is split (5,000 before); the setting is `Graph.SplitFolderItems`.
- The progress bar shows the share of the messages read, folder by folder, and the messages a second.
- HTML report: a **Read** column and a **Read** filter (all, read, unread).
- The log gives the speed of each mailbox and the slowest folders; `Folders.csv` has a **Seconds** column and the *Folders* tab a *Read in* column.

## 2.0.0 - 2026-10-07

Every message in the HTML report, and a CSV file Excel always opens. The format of the reports changes (MAJOR).

- The HTML report is new and holds every message (up to 500,000 by default, 20,000 before; setting `Report.HtmlMaxMessages`).
- The HTML report has combined filters, sorting by column, every recipient of a message, and *Export the view to CSV*.
- The CSV files have a new column, `RecipientCount`, after `Bcc`. Scripts that read the columns by position need a check.
- A cell of the CSV files never exceeds the size of an Excel cell; a long recipient list is cut after its last whole address, with the number of the others.
- The item ID stays in the CSV files only.

## 1.1.1 - 2026-10-07

Pages cut short by Exchange Online.

- A page cut short by Exchange Online no longer stops the run: the tool asks again, then with fewer messages, and leaves out only a message that cannot be returned.
- A message left out is counted in the folder detail and the run ends with warnings (exit code 2).
- Large folders: messages newer or older than the first dates read are included.
- The HTML report shows a notice when the browser cannot run its script.
- The log of a search from the window gives the recipients, the formats and the layout.

## 1.1.0 - 2026-10-07

The preview of the window, like Outlook.

- Folder view: the folders of each mailbox, its archive and Recoverable Items, with the messages found in each, newest first.
- Reading pane: subject, sender, recipients, dates, folder, message ID and the content of the selected message, read on demand with `Mail.Read` and written to the log. `Window.ReadBody = $false` turns it off.
- The reports never hold the content of a message.
- Preview of the first 10 messages of each folder (`Window.PreviewPerFolder`); *List* shows them in one sortable list.

## 1.0.0 - 2026-10-07

First version.

- The primary mailbox and the In-Place Archive are read through Microsoft Graph only: no Exchange Online PowerShell, no export.
- Recoverable Items of both with `-IncludeRecoverableItems`.
- Every message: mailbox, location, folder, dates, subject, sender, To, Cc, Bcc, Internet message ID, attachments, importance, read state and type.
- Filters: the period (`-Start`, `-End`), the subjects (`-Subject`), folders left out (`-ExcludeFolder`); `-SkipRecipients` for large folders.
- One mailbox (`-Mailbox`) or a list (`-MailboxFile`); a mailbox that cannot be read is reported with its reason.
- Large mailboxes are read in parallel and written to disk page by page, so the memory does not grow with the messages.
- Reports: CSV files, `Summary.json` and a self-contained HTML report; one report for all mailboxes, one per mailbox, or both.
- Console with progress and time left, exit codes 0, 1 and 2, a daily log, and a window (`-Gui`).
- Least permissions: `Mail.ReadBasic.All` and `User.Read.All`, with a certificate or a client secret.
