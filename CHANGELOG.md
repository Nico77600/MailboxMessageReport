# Changelog

All notable changes to Mailbox Message Report. Versions: MAJOR.MINOR.PATCH (developer guide, appendix D).

## 1.1.0 - 2026-10-07

The preview of the window, like Outlook.

- **Folder view**: the tree of each mailbox read — *Primary mailbox*, *Archive* and, when read, *Recoverable Items* under each — with the messages found per folder (a folder without any greyed, a folder not read in red); the messages of the folder selected (sender, subject, date, attachment), newest first; the first folder with messages selected after a search.
- **Reading pane**: subject, sender, To, Cc, Bcc, received and sent dates, location and folder, Internet message ID, and the **content** of the message selected, read then from Microsoft Graph as text (`GET /users/{id}/messages/{id}?$select=body`, cut at 200,000 characters), one message at a time, 300 ms after the selection, kept for the rest of the session. It needs the application permission `Mail.Read` (with `Mail.ReadBasic.All` the pane says so and shows the rest); each content read is written to the log (mailbox, location, folder, Internet message ID); `Window.ReadBody = $false` turns it off. The reports never hold the content.
- **Preview per folder**: the first 10 messages of each folder (`Window.PreviewPerFolder`, 0 to 1,000), 5,000 in all at most (`Window.PreviewMessages`, 0 to 50,000; 1,000 before); *List* shows the same messages in one sortable list.
- Window 1480 × 940. 33 Pester tests (the first messages of each folder, the folder tree, the folder view and the reading pane); the simulated Exchange answers a message by its ID with its body as text.
## 1.0.0 - 2026-10-07

First version.

- **The primary mailbox and the In-Place Archive through Microsoft Graph only**: the archive ID from `GET /beta/users/{id}/settings/exchange` (`inPlaceArchiveMailboxId` = `MBX:<ArchiveGuid>@<tenant ID>`, checked against `Get-EXOMailbox` in the lab), then the archive read like a mailbox (`/users/MBX:.../mailFolders`). Without `User.Read.All`, the `ArchiveGuid` of a CSV list. No Exchange Online PowerShell, no export.
- **Recoverable Items** of the primary mailbox and of the archive (`-IncludeRecoverableItems`): *Deletions*, *Purges*, *Versions*, *DiscoveryHolds*, *SubstrateHolds*, *Calendar Logging* — readable with `Mail.ReadBasic.All` (measured).
- **Every message**: mailbox, location (*Primary* / *Archive*), Recoverable Items, folder path and name, received and sent date and time (time zone of the report), subject, from, sender, To, Cc, Bcc, Internet message ID, attachments, importance, read, type, UTC date, item ID.
- **Filters of Exchange**: the period (`-Start`, `-End`, received date, end date included), the subjects (`-Subject`, `contains()`, any of them), written the way Exchange accepts them (the date first: *InefficientFilter* otherwise, measured); folders left out by path (`-ExcludeFolder`); `-SkipRecipients` for large folders of meeting messages.
- **Mailboxes**: `-Mailbox` (any alias, UPN) or `-MailboxFile` (text, or CSV with `PrimarySmtpAddress`... and `ArchiveGuid`); a mailbox on-premises or inactive (*MailboxNotEnabledForRESTAPI*), not found, or denied is reported with its reason.
- **Large mailboxes**: lists read page after page, 16 requests in flight, 4 at a time per mailbox (a primary mailbox and its archive are two mailboxes), the mailboxes served in turn; a folder of more than `Graph.SplitFolderItems` (5,000) items read in slices of its dates, side by side; 429 / 5xx retried after *Retry-After*.
- **Pages to disk**: each page is read as bytes and written by a compiled writer to the part file of its folder (never through a PowerShell string: about 100 ms per MB saved); the report merges them in order. The memory does not grow with the messages (200,000 synthetic messages: 4.6 s to write, 2.2 s for the report).
- **Reports**: `Messages.csv` (every message, cells safe for Excel), `Mailboxes.csv`, `Folders.csv`, `Summary.json`, a self-contained HTML report (messages, folders, mailboxes; the first `Report.HtmlMaxMessages` messages, 20,000 by default, up to 200,000 with `-HtmlMaxMessages`: 118 MB, open in about 3 s, 200 rows at a time, search after a pause in the typing, sort under a second); `-Layout Global`, `PerMailbox` (one report per mailbox in `Mailboxes\`, a summary that links to each) or `Both`. The HTML is written as it goes: the process stays under 500 MB for 200,000 messages.
- **Console** with steps, progress and time left, a summary card and exit codes (0, 1, 2); a daily log.
- **Window** (`-Gui`, Fluent theme of Windows 11 with PowerShell 7.5+): the search typed or loaded from a list, the mailboxes read, a preview of the first messages (`Window.PreviewMessages`), the progress with the time left; the run in a background runspace.
- **Least permissions**: `Mail.ReadBasic.All` and `User.Read.All` (measured with temporary applications); certificate (client assertion built by the tool) or client secret.
- 30 Pester tests on a simulated Exchange Online tenant; a measurement tool; the user guide and the developer guide (Markdown and HTML), the graphics of the GitHub page.
