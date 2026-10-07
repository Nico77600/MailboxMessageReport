#
#  Mailbox Message Report - configuration file
#  --------------------------------------------------------------------------
#  Author  : Nicolas Fabert
#  Version : 1.1.0
#
#  Read by Invoke-MailboxMessageReport.ps1 and by the window (-Gui). It is a PowerShell data file: text between
#  quotes, $true / $false, numbers, @( ) for lists and @{ } for groups of settings. Lines starting with # are
#  comments. Relative paths (.\reports, .\logs) are relative to the tool folder.
#  Every value is checked at start; all the problems are listed at once.
#
#  No secret here: the certificate stays in the Windows certificate store; a client secret is read from an
#  environment variable or typed at each run, and is never written.
#
@{
    # ---------------------------------------------------------------------
    # Tenant. Safety check: the tool stops if the token belongs to another tenant.
    # ---------------------------------------------------------------------
    Tenant = @{
        TenantId     = ''      # Microsoft Entra tenant ID (GUID) or its domain contoso.onmicrosoft.com
        Organization = ''      # optional, shown in the console and the report (contoso.onmicrosoft.com)
    }

    # ---------------------------------------------------------------------
    # Application registered in Microsoft Entra (developer guide, chapter 5). Application permissions of Microsoft
    # Graph, with admin consent:
    #   Mail.ReadBasic.All   required: folders and messages (subject, sender, recipients, dates, message ID) of the
    #                        primary mailbox, the archive and Recoverable Items - never the body. Mail.Read works too.
    #   User.Read.All        the user of each address (any alias) and the ID of its archive (settings/exchange).
    #                        Without it, the archive is read only when the list of mailboxes gives its ArchiveGuid.
    #   Certificate  : recommended. The certificate with its private key in Cert:\CurrentUser\My or
    #                  Cert:\LocalMachine\My of the account that runs the tool.
    #   ClientSecret : the secret is read from the environment variable ClientSecretVariable, or typed.
    # ---------------------------------------------------------------------
    Authentication = @{
        Mode                  = 'Certificate'          # Certificate | ClientSecret
        AppId                 = ''                     # application (client) ID
        CertificateThumbprint = ''                     # Certificate: thumbprint (40 hexadecimal characters)
        ClientSecretVariable  = 'MMR_CLIENT_SECRET'    # ClientSecret: name of the environment variable
    }

    # ---------------------------------------------------------------------
    # What is read (defaults of -Location, -IncludeRecoverableItems, -Start, -ExcludeFolder, -MailboxFile).
    #   Locations: 'Primary' (the mailbox), 'Archive' (the In-Place Archive), or both.
    #   RecoverableItems: also the folders of Recoverable Items (Deletions, Purges, Versions, DiscoveryHolds...):
    #     the items deleted from Deleted Items, or kept by a hold.
    #   Recipients: read To, Cc and Bcc. Exchange reads the recipients of each message for them: much slower on a
    #     large folder of meeting messages (lab: 28 s instead of 0.4 s for 1,000 messages). $false: the other columns only.
    #   PastDays: 0 = no start date (every message); else the messages received in the last N days.
    #   ExcludeFolders: folders left out, by path with wildcards, for example '\Junk Email', '\Inbox\Newsletters*',
    #     '\Recoverable Items\Purges'. The same paths in the primary mailbox and in the archive.
    #   MailboxFile: default list of mailboxes (one address per line, or a CSV file with PrimarySmtpAddress and
    #     optionally ArchiveGuid).
    # ---------------------------------------------------------------------
    Search = @{
        Locations        = @('Primary', 'Archive')
        RecoverableItems = $false
        Recipients       = $true
        PastDays         = 0
        ExcludeFolders   = @()
        MailboxFile      = ''
    }

    # ---------------------------------------------------------------------
    # Requests to Microsoft Graph: one request per page of messages, at most 4 at a time per mailbox (a primary
    # mailbox and its archive are two mailboxes for Exchange Online).
    #   SplitFolderItems: a folder with more items is read in slices of its received dates, 4 at a time (one slice
    #     per SplitFolderItems items, 16 at most); 0 = a folder is always read page after page.
    # ---------------------------------------------------------------------
    Graph = @{
        MaxConcurrency   = 16      # requests in flight (1-32)
        PageSize         = 250     # messages per page (10-1000)
        SplitFolderItems = 5000
        MaxRetries       = 6       # per request, for 429 / 5xx / no answer (the Retry-After delay is respected)
        TimeoutSeconds   = 120
    }

    # ---------------------------------------------------------------------
    # Report files (one sub-folder per run), written locally only.
    #   Layout: 'Global' one report for every mailbox, 'PerMailbox' one report per mailbox (folder Mailboxes\,
    #           plus a summary with a link to each one), 'Both'.
    #   HtmlMaxMessages: messages shown in a HTML report (each CSV file holds them all), 0 to 200,000. 200,000 makes a
    #     report of about 120 MB that a browser opens in about 3 s (search and sort under a second); 20,000: about 12 MB.
    # ---------------------------------------------------------------------
    Report = @{
        OutputPath      = '.\reports'
        FilePrefix      = 'MailboxMessageReport'
        Formats         = @('Csv', 'Html')   # a Summary.json file is always written as well
        Layout          = 'Global'           # Global | PerMailbox | Both
        HtmlMaxMessages = 20000
        CsvDelimiter    = ';'                # ';' opens directly in Excel with French regional settings
        TimeZone        = ''                 # dates typed and shown: '' = the time zone of Windows, or Europe/Paris...
    }

    # ---------------------------------------------------------------------
    # The window (-Gui): a preview of the messages - the first PreviewPerFolder messages of each folder (newest first),
    # PreviewMessages in all at most - shown as a list, or as folders with a reading pane like Outlook. The report
    # holds every message.
    #   ReadBody: a message selected in the folder view shows its content, read then from Microsoft Graph as text.
    #     It needs the application permission Mail.Read (Mail.ReadBasic.All does not give the body); each read is
    #     written to the log. $false: the window never reads a body.
    # ---------------------------------------------------------------------
    Window = @{
        PreviewMessages  = 5000
        PreviewPerFolder = 10
        ReadBody         = $true
    }

    # ---------------------------------------------------------------------
    # Log files (one per day, no token, no colour).
    # ---------------------------------------------------------------------
    Logging = @{
        Path          = '.\logs'
        RetentionDays = 30
    }
}
