---
title: Mailbox Message Report
subtitle: User guide
version: 2.0.0
author: Nicolas Fabert
updated: 2026-10-07
---

# Mailbox Message Report — User guide

> What you need before the first run, then one command per everyday question: **which messages are in this archive?**, **did these people receive this message, and where is it now?**, **what arrived in this period?**, **a report for every mailbox of a team**. How the archive is read through Microsoft Graph, the rights in detail, the configuration, the report and the internals are in the [developer guide](MailboxMessageReport-Guide.md).

> [!IMPORTANT]
> Files downloaded from the Internet may be blocked by Windows and fail to run. Before using this project, unblock every file in the downloaded folder:
>
> ```powershell
> Get-ChildItem "C:\Chemin\Du\Dossier" -Recurse -File -Force | Unblock-File
> ```
>
> Replace the example path with the folder where you downloaded or extracted this project.

```cards
checklist | Prerequisites | Chapter 1: PowerShell, the application and its certificate, then the one-time setup.
terminal | Everyday use | Chapter 2: one command per question, in the console or in the window.
filter | Choose the messages | Chapter 3: which mailboxes, where, which period, which subjects.
file | Results | Chapter 4: where the report is written, what it says, exit codes, the usual messages.
```

<!-- icon: checklist -->
## 1. Prerequisites

| Item | Requirement |
|---|---|
| Exchange | **Exchange Online** only: the primary mailbox, its In-Place Archive and Recoverable Items. Not the auxiliary archives of an auto-expanding archive. |
| PowerShell | **7.4 or later** (`pwsh`) — a portable zip is enough. With 7.5 and later the window has the Fluent look of Windows 11. |
| Windows | Windows 10 / 11, Windows Server 2016 to 2025. The window needs a desktop session; the command line runs anywhere (scheduled task, SSH). |
| Application | An **application registered in Microsoft Entra**, with the application permissions `Mail.ReadBasic.All` and `User.Read.All` (admin consent) and a **certificate** whose private key is in the certificate store of the account that runs the tool: [developer guide, chapter 5](MailboxMessageReport-Guide.md#5-application). |
| Your account | **No Exchange or Entra role** to run it: the application signs in with its certificate. No Exchange Online PowerShell. |
| Network | HTTPS to `login.microsoftonline.com` and `graph.microsoft.com`. |

> [!WARNING]
> `Mail.ReadBasic.All` as an application permission reaches **every mailbox** of the tenant (never the body or the attachments of a message). Keep the certificate like an administrator password (private key not exportable, on the computer that runs the tool), or limit the application to some mailboxes with **RBAC for Applications** ([developer guide, chapter 5](MailboxMessageReport-Guide.md#5-application)).

### 1.1 One-time setup

```steps
Copy the tool | Unblock the files, then copy the folder, for example to `C:\Tools\MailboxMessageReport`. No installer.
Certificate | On the computer that runs the tool, as the account that runs it: the commands below. Keep the thumbprint.
Application | Microsoft Entra > *App registrations*: `Mail.ReadBasic.All` and `User.Read.All`, **Grant admin consent**, then upload the `.cer` file ([developer guide, chapter 5](MailboxMessageReport-Guide.md#5-application)).
Configure | `notepad .\config\MailboxMessageReport.config.psd1`: `Tenant.TenantId`, `Tenant.Organization`, `Authentication.AppId`, `Authentication.CertificateThumbprint`.
Check | `.\Invoke-MailboxMessageReport.ps1 -Mailbox <your address> -Start <yesterday>`: step 1 shows the application and the permissions of its token; nothing is ever changed.
```

```powershell
# On the computer that runs the tool, as the account that runs it
$cert = New-SelfSignedCertificate -Subject 'CN=MailboxMessageReport' -CertStoreLocation Cert:\CurrentUser\My `
    -KeyExportPolicy NonExportable -KeySpec Signature -KeyAlgorithm RSA -KeyLength 2048 -NotAfter (Get-Date).AddYears(2)
Export-Certificate -Cert $cert -FilePath .\MailboxMessageReport.cer      # upload this file to the application
$cert.Thumbprint                                                         # Authentication.CertificateThumbprint
```

<!-- icon: terminal -->
## 2. Everyday use

Run the commands from the tool folder, in PowerShell 7. Each command reads the mailboxes and writes a report; **nothing is ever changed** in a mailbox. `-Gui` does the same in a window (2.7).

### 2.1 Which messages are in this archive?

```powershell
# The archive only
.\Invoke-MailboxMessageReport.ps1 -Mailbox megan.bowen@contoso.com -Location Archive

# The primary mailbox and the archive (default)
.\Invoke-MailboxMessageReport.ps1 -Mailbox megan.bowen@contoso.com
```

Open the HTML report: the **Messages** tab lists every message, with the *Location* (*Primary* or *Archive*) and the folder path; the **Folders** tab gives the number of items of each folder.

### 2.2 Did these people receive this message, and where is it now?

```powershell
.\Invoke-MailboxMessageReport.ps1 -Mailbox megan.bowen@contoso.com, alex.wilber@contoso.com -Subject 'Contrat Alpha' -IncludeRecoverableItems
```

The subject *contains* the text, any case; several subjects: `-Subject 'Contrat Alpha', 'Projet Alpha'` (any of them). With `-IncludeRecoverableItems`, the messages deleted from *Deleted Items* (or kept by a hold) are listed too, with the path `\Recoverable Items\Deletions`. The *InternetMessageId* column is the same for one message in every mailbox: sort on it to see who has it, where.

### 2.3 What arrived in a period?

```powershell
.\Invoke-MailboxMessageReport.ps1 -Mailbox megan.bowen@contoso.com -Start 2026-09-01 -End 2026-09-30
```

The received date, in the time zone of the configuration (`Report.TimeZone`, Windows by default); the end date is included.

### 2.4 A report for every mailbox of a team

```powershell
.\Invoke-MailboxMessageReport.ps1 -MailboxFile .\team.csv -Start 2026-01-01 -Layout Both
```

A text file with one address per line, or a CSV file (`PrimarySmtpAddress`, `UserPrincipalName`, `Mail`...). `-Layout Global` (default) writes one report for all the mailboxes, `PerMailbox` one report per mailbox (folder `Mailboxes\`, with a summary that links to each one), `Both` the two.

### 2.5 The archive without User.Read.All

```powershell
# Once, by an Exchange administrator: the ArchiveGuid of each mailbox
Get-EXOMailbox -ResultSize Unlimited -Archive -Properties ArchiveGuid | Select-Object PrimarySmtpAddress, ArchiveGuid | Export-Csv .\archives.csv -NoTypeInformation

.\Invoke-MailboxMessageReport.ps1 -MailboxFile .\archives.csv
```

With `User.Read.All`, the tool finds the archive itself in Microsoft Graph. Without it, the `ArchiveGuid` column of the list gives it.

### 2.6 A very large mailbox

```powershell
# The other columns first: much faster on large folders of meeting messages
.\Invoke-MailboxMessageReport.ps1 -Mailbox room-paris-01@contoso.com -SkipRecipients

# Leave out the folders you do not need
.\Invoke-MailboxMessageReport.ps1 -Mailbox megan.bowen@contoso.com -ExcludeFolder '\Junk Email', '\Deleted Items'
```

Exchange reads the recipients (To, Cc, Bcc) of each message one by one: on a folder of tens of thousands of meeting messages, 1,000 messages take about 30 s with them and half a second without (lab: 39,434 messages in 7 min 27 s with them, 43 s without). A folder of more than 5,000 items is read in slices of its dates, 4 at a time. The progress line gives the time left.

### 2.7 In the window

```powershell
.\Invoke-MailboxMessageReport.ps1 -Gui
```

![The window after a search](images/gui-search-light.png)

```steps
Mailboxes | Type the addresses (one per line), or *Load a list...* (text or CSV, with `ArchiveGuid` if you have it).
Messages | The period (empty: no limit), the subjects (one per line), the recipients; *Where*: primary mailbox, archive, Recoverable Items, folders left out; the report: CSV, HTML, global or per mailbox.
Read | *Read the messages*: the progress bar gives the step, the part done and the time left; *Stop* ends the run.
Look | The mailboxes read, and a preview: the first 10 messages of each folder. *Folders*: the folders of the mailbox and of its archive like in Outlook; select a folder, then a message: the reading pane shows its recipients, its dates, where it is and its content. *List*: every message of the preview in one list.
Report | *Open the report* (every message), *Open the CSV*, *Open the folder*.
```

> [!NOTE]
> The content of a message is read only when you select it, and only if the application has the permission `Mail.Read` (with `Mail.ReadBasic.All` the reading pane shows the rest). Each content read is written to the log. The reports never hold the content.

<!-- icon: filter -->
## 3. Choose the messages

| Parameter | Values | Example |
|---|---|---|
| `-Mailbox` | One or more addresses (any alias) or UPNs | `-Mailbox megan.bowen@contoso.com` |
| `-MailboxFile` | A list: text (one per line) or CSV (`ArchiveGuid` optional) | `-MailboxFile .\team.csv` |
| `-Location` | `Primary`, `Archive` or both (default) | `-Location Archive` |
| `-IncludeRecoverableItems` | Also Recoverable Items (deleted items, holds) | `-IncludeRecoverableItems` |
| `-Start` · `-End` | The received date (end included). Default: no limit | `-Start 2019-01-01 -End 2019-12-31` |
| `-Subject` | The subject contains this text; several: any of them | `-Subject 'Invoice', 'Facture'` |
| `-ExcludeFolder` | Folders left out, by path (`*` allowed) | `-ExcludeFolder '\Junk Email'` |
| `-SkipRecipients` | Without To, Cc, Bcc: faster | `-SkipRecipients` |
| `-Layout` | `Global` (default), `PerMailbox`, `Both` | `-Layout PerMailbox` |
| `-Format` | `Csv`, `Html` or both (default) | `-Format Csv` |
| `-HtmlMaxMessages` | The most messages of the HTML report (default 500,000, up to 2,000,000) | `-HtmlMaxMessages 2000000` |
| `-Gui` | The window | `-Gui` |

<!-- icon: file -->
## 4. Results

Each run writes a new folder under `reports\` (`MailboxMessageReport_20261007-101500`). The console shows it at the end:

- **`MailboxMessageReport.html`** — the report: self-contained, it can be sent alone. Tiles (messages in the primary mailbox, the archive, Recoverable Items), the search, and the *Messages*, *Folders* and *Mailboxes* tabs. **Every message** is in it, compressed (up to 500,000, `Report.HtmlMaxMessages`; 1,000,000 messages: about 32 MB, open in 4 s):
  - filters together: every field, mailbox, location, folder, sender, subject, recipient (To, Cc, Bcc), message ID, received dates, a minimum of recipients, with attachments;
  - a click on a column sorts it; a click on a message gives every detail and **every recipient** (To, Cc, Bcc, with *Copy*);
  - *Export the view to CSV*: the messages filtered, in the order shown.
- **`MailboxMessageReport-Messages.csv`** — every message: mailbox, location, Recoverable Items, folder path and name, received, sent, subject, from, sender, to, cc, bcc, **number of recipients**, Internet message ID, attachments, importance, read, type. Separator `;`, opens directly in Excel. A list of thousands of recipients is cut to fit an Excel cell, with the number of the others (`… (+8,739 more)`): the HTML report has them all.
- **`MailboxMessageReport-Mailboxes.csv`**, **`-Folders.csv`** — one row per mailbox (archive, messages per location, why a mailbox was not read), one row per folder.
- **`MailboxMessageReport-Summary.json`** — the whole result, for scripts.
- **`Mailboxes\`** — with `-Layout PerMailbox` or `Both`: the CSV and HTML report of each mailbox.

| Mailbox status | Meaning |
|---|---|
| **Read** | Every folder asked for was read. |
| **Partial** | A folder (or the archive) could not be read: see *Notes*, and the *Folders* tab. |
| **Not read** | Mailbox on-premises or inactive, not a mailbox of the tenant, or access denied. |

| Exit code | Meaning |
|---|---|
| `0` | Completed. |
| `2` | Finished with warnings: a mailbox or a folder not read, an archive not known. The report says which. |
| `1` | Failed: read the error in red, and the log of the day in `logs\`. |

| Message | What to do |
|---|---|
| `AADSTS700016` · `AADSTS700027` | Application ID not in the tenant · certificate not uploaded to the application, or another one. |
| *Certificate ... not found* | Import the `.pfx` (not the `.cer`) for the account that runs the tool, in `Cert:\CurrentUser\My` or `LocalMachine\My`. |
| *no application permission Mail.ReadBasic.All* | Permission missing, or admin consent not granted (`Mail.Read` counts as well). |
| *archive not known* | `User.Read.All` missing: add it, or give the `ArchiveGuid` in a CSV list (2.5). |
| *mailbox inactive, soft-deleted or on-premises* | Graph cannot open this mailbox (hybrid: the mailbox is on-premises). |
| A folder of the archive *Failed* (*auxiliary archive*) | Auto-expanding archive: that folder is out of reach of the tool. |
| *access denied* | The application is limited to some mailboxes (RBAC for Applications). |
| *Graph answered a page cut short* in the log | Exchange Online ended a page early: the tool asks it again with fewer messages. *N message(s) left out* in the summary: Graph could not return them; the *Folders* tab says where. Many of them: `-SkipRecipients`. |
| The HTML report shows no message | With `-Layout PerMailbox`, open the report of each mailbox (*Mailboxes* tab). Otherwise open the report in Microsoft Edge or Google Chrome (not in Internet Explorer mode, not in the preview of OneDrive); the CSV file holds every message. |

Anything else: [developer guide, Appendix A — Troubleshooting](MailboxMessageReport-Guide.md#appendix-a---troubleshooting); every column of the report: [developer guide, chapter 4](MailboxMessageReport-Guide.md#4-what-is-read).
