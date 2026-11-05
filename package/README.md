# Mailbox Message Report

Mailbox Message Report lists the messages of the primary mailbox and the In-Place Archive (the second, hidden mailbox of a user) of one mailbox or thousands, with the date, subject, sender, recipients, message ID and folder of each message and where it is, through Microsoft Graph only. It only reads: nothing is ever changed in a mailbox.

This folder is the release ZIP, extracted. It holds everything needed to run the tool: `Invoke-MailboxMessageReport.ps1`, the module, the configuration, the report template and the two guides (open them in your browser).

> [!IMPORTANT]
> Files downloaded from the Internet may be blocked by Windows. Unblock them once, from this folder, in PowerShell 7:
>
> ```powershell
> Get-ChildItem . -Recurse -File -Force | Unblock-File
> ```
>
> Only `config\MailboxMessageReport.config.psd1`, `reports\` and `logs\` may change inside the tool folder; the tool refuses to start when another file is added, changed or missing. Keep your own files (mailbox lists, the certificate export) in a separate folder, for example `C:\Tools\MailboxMessageReport-data`.

## Get started

1. **Keep your own files outside the tool folder** (mailbox lists, the certificate export), for example in `C:\Tools\MailboxMessageReport-data`, and give their absolute path.
2. **Open PowerShell 7 in this folder.** In File Explorer, open this folder, click the address bar, type `pwsh` and press Enter. Run `$PSVersionTable.PSVersion`: you should see 7.4 or higher. If `pwsh` is not recognized or the version starts with 5, install PowerShell 7 (`winget install Microsoft.PowerShell`) and open a new window: Windows PowerShell 5.1 is not enough.
3. **Register the application** in Microsoft Entra and note its IDs: follow chapter 1 of the user guide, `docs\MailboxMessageReport-UserGuide.html`. It lists, for each step, what you should see and what to do if not (the table of messages is in its chapter 4).
4. **Fill in the configuration**: `notepad .\config\MailboxMessageReport.config.psd1` (tenant ID, application ID, certificate thumbprint).
5. **Run a first report**: `.\Invoke-MailboxMessageReport.ps1 -Mailbox megan.bowen@contoso.com` (replace the address with a mailbox of your tenant). You should see the steps in the console, then a new folder under `reports\` with `MailboxMessageReport.html` and the CSV files. Or open the window: `.\Invoke-MailboxMessageReport.ps1 -Gui`. If it fails, the error is shown in red: look it up in chapter 4 of the user guide.

More everyday commands:

```powershell
.\Invoke-MailboxMessageReport.ps1 -Mailbox megan.bowen@contoso.com -Location Archive -Start 2019-01-01 -End 2019-12-31
.\Invoke-MailboxMessageReport.ps1 -MailboxFile C:\Tools\MailboxMessageReport-data\team.csv -Subject 'Contract Alpha' -IncludeRecoverableItems -Layout Both
.\Invoke-MailboxMessageReport.ps1 -Mailbox room-paris-01@contoso.com -SkipRecipients
```

## Check the download

The PowerShell files carry the owner's timestamped Authenticode signature (Nicolas Fabert), and the tool checks all its files against a signed list before it starts. To check by hand:

```powershell
Get-AuthenticodeSignature $PWD\Invoke-MailboxMessageReport.ps1 | Format-List Status, SignerCertificate
```

You should see `Status : Valid` and the signer Nicolas Fabert. The SHA256 of the ZIP is on the release page. If the tool answers *Package integrity*, a file was changed or is missing: download the ZIP again from the Releases page and extract it again.

## Requirements
- Exchange Online primary mailbox, In-Place Archive and Recoverable Items.
- PowerShell 7.4 or later.
- Windows 10 / 11 or Windows Server 2016 to 2025.
- A Microsoft Entra application with the application permissions Mail.ReadBasic.All and User.Read.All (admin consent).
- A certificate whose private key is in the Windows store of the account that runs the tool.
- HTTPS to login.microsoftonline.com and graph.microsoft.com.

## Content
| Item | Role |
|---|---|
| Invoke-MailboxMessageReport.ps1 | Entry script for the window and command line. |
| MailboxMessageReport.psd1 | Module manifest. |
| MailboxMessageReport.psm1 | Module loader. |
| PackageIntegrity.psd1 | Signed list of the files of this package. Do not edit. |
| README.md | This file. |
| config | Configuration file to edit. |
| docs | The user guide and the developer guide, as HTML files. |
| src | PowerShell source files and native helper source. |
| templates | HTML report template. |
| CHANGELOG.md | What changed in each version. |
| LICENSE | MIT license. |
| THIRD-PARTY-NOTICES.md | Third-party notices. |

## Documentation
- User guide: `docs\MailboxMessageReport-UserGuide.html` - a single file to open in your browser
- Developer guide: `docs\MailboxMessageReport-Guide.html`

Project page and releases: https://github.com/Nico77600/MailboxMessageReport

License: see the LICENSE file (MIT).

## Disclaimer

This Script is a Personal project.
It's provided "AS-IS". It's not an official Microsoft product so no support can be expected from Microsoft.

As any scripts you must read carefully the documentation and test it first in a Test environment before any run in Production.

The archive is read through an ID that Microsoft Graph gives (a beta endpoint) and opens, but that its documentation does not describe for the archive: test it in your environment before production use. The reports contain personal data (addresses, subjects): store and share them accordingly.
