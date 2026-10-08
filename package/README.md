# Mailbox Message Report

Mailbox Message Report lists the messages of the primary mailbox and the In-Place Archive of one mailbox or thousands, with the date, subject, sender, recipients, message ID and folder of each message and where it is, through Microsoft Graph only.

This folder contains everything needed to run the tool: Invoke-MailboxMessageReport.ps1, the module, the configuration, the report template and the guides. Tests and build tools stay outside it, in the repository.

> [!IMPORTANT]
> Files downloaded from the Internet may be blocked by Windows. Unblock them once, from this folder:
>
> ```powershell
> Get-ChildItem . -Recurse -File | Unblock-File
> ```

## Requirements
- Exchange Online primary mailbox, In-Place Archive and Recoverable Items.
- PowerShell 7.4 or later.
- Windows 10 / 11 or Windows Server 2016 to 2025.
- Microsoft Entra application permissions: Mail.ReadBasic.All and User.Read.All.
- Certificate private key in the Windows store of the account that runs the tool.
- HTTPS to login.microsoftonline.com and graph.microsoft.com.

## Quick start
```powershell
notepad .\config\MailboxMessageReport.config.psd1          # tenant, application, certificate thumbprint

.\Invoke-MailboxMessageReport.ps1 -Gui                     # the window

# Or the command line: nothing is ever changed in a mailbox
.\Invoke-MailboxMessageReport.ps1 -Mailbox megan.bowen@contoso.com                                   # primary mailbox and archive
.\Invoke-MailboxMessageReport.ps1 -Mailbox megan.bowen@contoso.com -Location Archive -Start 2019-01-01 -End 2019-12-31
.\Invoke-MailboxMessageReport.ps1 -MailboxFile .\team.csv -Subject 'Contrat Alpha' -IncludeRecoverableItems -Layout Both
.\Invoke-MailboxMessageReport.ps1 -Mailbox room-paris-01@contoso.com -SkipRecipients                # very large folders, fast
```

## Content
| Item | Role |
|---|---|
| Invoke-MailboxMessageReport.ps1 | Entry script for the window and command line. |
| MailboxMessageReport.psd1 | Module manifest. |
| MailboxMessageReport.psm1 | Module loader. |
| config | Delivered configuration template. |
| docs | User and developer guides in Markdown and HTML, with images. |
| src | PowerShell source files and native helper source. |
| templates | HTML report template. |
| LICENSE | MIT license. |
| THIRD-PARTY-NOTICES.md | Third-party notices. |

## Documentation
- [User guide](docs/MailboxMessageReport-UserGuide.md) - also `docs/MailboxMessageReport-UserGuide.html`, a single file to open locally
- [Developer guide](docs/MailboxMessageReport-Guide.md) - also `docs/MailboxMessageReport-Guide.html`

Project page, releases and change log: https://github.com/Nico77600/MailboxMessageReport

License: [MIT](LICENSE).
