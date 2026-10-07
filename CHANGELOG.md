# Changelog

All notable changes to Mailbox Message Report. Versions: MAJOR.MINOR.PATCH (developer guide, appendix D).

## 2.1.0 - 2026-10-07

Large folders read 4 requests at a time to the end, and a progress bar that shows the messages read.

- **Slices of the same number of messages**: a folder of more than `Graph.SplitFolderItems` messages (now **2,000**, 5,000 before) is cut at the dates of the messages at 1/n, 2/n... of the folder (`$top=1&$count=true`, then `$top=1&$skip=k·N/n`), 16 slices at most. 2.0 cut the period in equal parts: a room's *Deleted Items* had 39,372 of its 39,419 messages in 3 days of 6 months, one slice held almost all of them, read one page after the other.
- **Slices cut again while read**: when a mailbox has free request slots (Exchange answers 4 at a time per mailbox), the list with the longest period left is cut in two at the middle (the list in course stops there, a new slice reads the older half; 64 per folder at most; each cut in the log). Every message is read once, newest first.
- Lab, that folder, 6 months, with the recipients: **14 min 06 s → 3 min 08 s** (214 messages a second, 3.9 requests at a time on average instead of 1).
- **Progress bar**: the share of the messages read, folder by folder (the number of messages of the filter of each list, `$count` on its first page). 2.0 counted the empty slices of a folder as read: the bar stayed at 100 %. The progress line gives the messages a second.
- **HTML report**: a **Read** column in the table of the messages (*No* in colour, the subject of an unread message in bold) and a **Read** filter (all, read, unread), with the others; sortable like every column.
- **Speed in the log**: the time of a page (average, longest), the requests in flight on average, the messages a second of the run and of each mailbox, the 5 slowest folders. New column **Seconds** in `Folders.csv` and *Read in* in the *Folders* tab of the HTML report.
- 38 Pester tests (cuts at k·N/n, a folder whose messages came in a few days, a slice cut again while read, the progress); the simulated Exchange answers `$count=true`.
## 2.0.0 - 2026-10-07

Every message in the HTML report, and a CSV file Excel always opens. The format of the reports changes (MAJOR).

- **HTML report like the one of Purview DLP Report**: every message (up to `Report.HtmlMaxMessages`, now 500,000 by default, 0 to 2,000,000; 20,000 before), in compressed blocks of 20,000 (JSON, gzip, base64; folders, subjects, addresses, names and types once in dictionaries), decompressed by the page into typed arrays, in a table that draws only the visible rows. 200,000 messages: 6.5 MB, open in 1.1 s; 1,000,000: 32 MB, 4.1 s, 305 MB in the page (1.x: 200,000 messages in 118 MB, 200 rows at a time).
- **Filters** together: every field, mailbox, location, folder, sender, subject, recipient (To, Cc, Bcc), Internet message ID, received dates, minimum of recipients, with attachments; **sort** by a click on a column; **every recipient** of a message in its detail (To, Cc, Bcc with *Copy*); **Export the view to CSV**.
- **CSV safe for Excel**: a cell is never longer than an Excel cell (32,767 characters). A list of recipients is cut after its last whole address, with the number of the others (`… (+8,739 more)`). A message sent to 10,000 people broke the row in Excel (4 rows × 8,750 columns, measured): one row of 23 columns now.
- **New column `RecipientCount`** (To + Cc + Bcc; empty with `-SkipRecipients`) in the CSV files, after `Bcc`.
- `MailboxMessageReportNative.HtmlReport` replaces `Merge.WriteHtml`; the columns are read by name (`Columns.I*`). The item ID stays in the CSV files only.
- 36 Pester tests (the blocks of the HTML report decoded as the page does, a message of 10,000 recipients); the measurement tool writes realistic pages (every message ID unique, 300 people).
## 1.1.1 - 2026-10-07

Pages cut short by Exchange Online.

- **A page cut short no longer stops the run.** Exchange Online sometimes answers a page of messages with a status 200 and a JSON that ends early (seen on a customer tenant on a slice of a large folder: `Exception calling "AddPage"... Expected depth to be zero at the end of the JSON payload`, and the run stopped without a report). The page is now asked again once; then with half as many messages, down to one (the next pages grow back to `Graph.PageSize`); then that one message without its sender and recipients; at last it is left out and the list goes on. Each step is logged (*Graph answered a page cut short*); a message left out or without sender and recipients is counted in the *Detail* of its folder and the run finishes with warnings (exit code 2). A list of folders cut short is asked again (`Graph.MaxRetries`), never kept incomplete.
- **Large folders**: the newest slice has no end and the oldest no start (but those of the period asked): a message newer or older than the two dates read first is read as well (checked in the lab: `lt` or `ge` alone accepted with `$orderby`; 39,434 messages, each once).
- **HTML report**: a notice when the browser cannot run its script (scripts turned off, an old engine) instead of an empty page.
- The log of a search from the window gives the recipients asked, the formats and the layout of the report.
- 35 Pester tests (a page cut short asked again smaller, a message Graph cannot return left out, larger pages cut short, a page and a list of folders cut once).
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
