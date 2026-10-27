# Usage Notes

- Place input CSVs into `data/input/`.
- Results and reports will be written to `data/output/`.
- Old runs can be moved to `data/archive/`.

Example:

```pwsh
pwsh -File src/Exchange/Export-SharedCalendar.ps1 -SharedMailboxEmail 'shared@contoso.com'
```

New options for mailbox selection (interactive `-ListMailboxes` or non-interactive):

- `-MailboxType` accepts `Shared` (default), `User`, or `All`.
- `-InternalOnly` restricts to mailboxes in your tenant domain (defaults to the signed-in account domain).
- `-ExternalOnly` restricts to mailboxes outside your tenant domain.
- `-Domains <domain[]>` lets you explicitly provide domain(s) used for internal/external filtering.

Examples:

```pwsh
# List shared mailboxes interactively
pwsh -File src/Exchange/Export-SharedCalendar.ps1 -ListMailboxes -MailboxType Shared

# Export all internal user mailboxes (non-interactive)
pwsh -File src/Exchange/Export-SharedCalendar.ps1 -MailboxType User -InternalOnly -WhatIf

# Orchestrate bulk exports across mailboxes (use -WhatIf to preview)
pwsh -File src/Exchange/Export-Mailboxes.ps1 -MailboxType User -InternalOnly -WhatIf

# Limit the sample size while testing
pwsh -File src/Exchange/Export-Mailboxes.ps1 -MailboxType Shared -Max 5 -WhatIf -ConfirmRun
```
Keep sensitive information (credentials, tokens) out of the repo — use environment variables or a secrets store.

---

Logging helpers:

- `Write-LogMessage` writes a timestamped entry to console and a log file. Use `-Level` set to `INFO`, `WARN`, or `ERROR`.
- `Set-LogFile -Path <file>` sets the helper's log file so subsequent `Write-LogMessage` calls append there.
- `Clear-LogFile` clears the configured file.

Example:

```pwsh
. "$PSScriptRoot/src/Utils/LoggingHelpers.ps1"  # or: Import-Module ./ExchangeMailboxToolkit
Set-LogFile -Path "$PSScriptRoot/logs/my-run.log"
Write-LogMessage -Message 'Starting run' -Level INFO
```
