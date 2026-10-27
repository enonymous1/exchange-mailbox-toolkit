# exchange-mailbox-toolkit

PowerShell 7.2+ tools for Exchange Online / Microsoft 365 mailbox administration: enforce a mailbox folder structure, export mailboxes, calendars and messages, verify folder structures, and move messages between folders, with Graph retry logic and audit logging.

This is v2.0.0, a modernization of a working toolkit that grew script by script. The Graph logic is unchanged in intent; what changed is how it is packaged, loaded, tested, and gated before anything destructive runs.

Formerly named `mp-workingDir` / `MPWorkingDir`, with `MP`-prefixed function names. The repo, module, and function names were generalized before the first public push; scripts calling the old names need updating — see the mapping in the CHANGELOG. The one exception is `Get-Mailboxes`, kept as a deprecated shim that forwards to `Get-GraphMailboxes`.

## What it does

- **Folder structure** — `Ensure_Directory_Structure.ps1` exports a mailbox's folder hierarchy, creates folders from a specification file, and removes folders not in the specification after moving their messages to Inbox. System folders (Inbox, Sent Items, etc.) are protected.
- **Exports** — mailboxes (`Export-Mailboxes.ps1`), shared calendars (`Export-SharedCalendar.ps1`), mailbox messages and attachments (`Export-MailboxMessages.ps1`), and mailbox time zones / regional settings (`Export-Mailbox-Timezones.ps1`).
- **Inspection and repair** — `Inspect-MailboxFolders.ps1`, `Verify-Mailbox-Folder-Structure.ps1`, `Check_For_Messages.ps1`, and `Get-GraphMailboxes.ps1` (shared/user/all, internal/external filtering).
- **Moves** — `Import-MoveMessages.ps1` moves messages from a CSV map, and `Sort-Inbox.ps1` is an interactive per-message sorter.
- **Filesystem** — `FileList_Generator.ps1` inventories a directory tree to CSV.

## Quick start

Prerequisites: PowerShell 7.2+, plus `Microsoft.Graph` (2.0.0+) and `ExchangeOnlineManagement` (3.0.0+) for any script that touches Graph or Exchange. Those modules are checked at runtime by `Connect-ExchangeEnvironment`, not forced at import time, so the pure helpers and the tests run without them.

```powershell
# Use the module (recommended for the helpers)
Import-Module ./ExchangeMailboxToolkit
Connect-ExchangeEnvironment -GraphScopes 'User.Read.All','Mail.Read.Shared' -Verbose

# Or run a script directly (scripts dot-source the same helpers)
pwsh -File src/Exchange/Ensure_Directory_Structure.ps1 -MailboxEmail user@contoso.com -DryRun
pwsh -File src/Exchange/Check_For_Messages.ps1 -CsvPath data/input/messages.csv -IdFilePath data/input/ids.txt -Save
pwsh -File src/Filesystem/FileList_Generator.ps1 -Path ./src -OutputFile data/output/file_list.csv
```

## Safety first

- **Read-only first.** Export, inspect, and verify scripts change nothing. For the write scripts, run with `-DryRun` where it exists, and `-WhatIf` on the scripts that support it.
- **Destructive operations are gated.** `Ensure_Directory_Structure.ps1` (folder removal), `Remove-MailboxFolders.ps1` (folder deletion), and `Import-MoveMessages.ps1` (message moves) all declare `SupportsShouldProcess`, so `-WhatIf` and `-Confirm` work. `Ensure` also moves messages to Inbox before deleting a folder, and never deletes a protected system folder.
- **Interactive scripts are labelled as such.** `Sort-Inbox.ps1` is an interactive sorter by design, and `Export-SharedCalendar.ps1` can browse mailboxes interactively. Pass `-SharedMailboxEmail` to run the calendar export non-interactively. `Export-Mailboxes.ps1 -ConfirmRun` is an opt-in confirmation gate for bulk exports.
- **Outputs stay out of git.** `data/` and `logs/` and `*.csv` / `*.json` / `*.pst` exports are git-ignored, so mailbox data does not end up in a commit.

## Project structure

```text
exchange-mailbox-toolkit/
├── ExchangeMailboxToolkit/          # Module: ExchangeMailboxToolkit.psd1 / .psm1 (v2.0.0)
├── src/
│   ├── Exchange/          # Exchange Online / Graph scripts
│   ├── Filesystem/        # Local filesystem utilities
│   └── Utils/             # Graph, logging, and Connect-ExchangeEnvironment helpers
├── examples/              # Debug-Attachment.ps1 (diagnostic only, changes nothing)
├── tests/                 # Pester tests (ExchangeMailboxToolkit.Tests.ps1)
├── docs/                  # Usage, quick reference, and execution history
├── .github/workflows/     # CI: manifest test, PSScriptAnalyzer, Pester
├── CHANGELOG.md
└── README.md
```

## Tests and CI

```powershell
# Syntax-check every script (no Graph session needed)
# Pester suite (needs Pester 5+; CI installs it)
Invoke-Pester -Path ./tests
```

CI runs on Windows and Ubuntu with `pwsh`: manifest validation, PSScriptAnalyzer, and the Pester suite. Live Graph/Exchange behaviour is not tested in CI — that gate is a real tenant run, scoped to one mailbox first (see `docs/usage.md`).

## What changed in v2.0.0

- **Fixed two scripts that did not parse at all in v1**, despite v1's docs claiming syntax validation passed: `Ensure_Directory_Structure.ps1` was missing two closing braces in its folder-creation loop, and `Get-Mailboxes.ps1` ended its header comment with `>` instead of `#>`, so the whole file was one unterminated comment.
- **Fixed the loader.** `Connect-ExchangeEnvironment.ps1` did not exist in v1 even though the README, Copilot instructions, and two scripts referenced it. It now exists as a loader for the function in `Initialize-Environment.ps1`, and importing the helpers no longer triggers a sign-in on its own.
- **Proper module.** `ExchangeMailboxToolkit.psd1` / `.psm1`, version 2.0.0, PowerShell 7.2, 13 exported functions.
- **Cross-platform paths.** Filesystem paths that used Windows backslashes in `Join-Path` and dot-source strings now use forward slashes / `Join-Path` segments, so the scripts run under `pwsh` on macOS/Linux too. Mailbox folder paths still use backslashes internally — that is the mailbox hierarchy separator, not a filesystem path, and it was left alone.
- **Non-interactive where it matters.** `Check_For_Messages.ps1` is now parameter-driven (and its id parser now handles files that mix commas and newlines, which v1 silently under-counted), and `Ensure_Directory_Structure.ps1` takes `-MailboxEmail` as mandatory and gates folder removal with `ShouldProcess` instead of `Read-Host`.
- **Implemented the stub.** `Export-Mailbox-Timezones.ps1` was a placeholder that printed a message; it now exports real regional configuration via Exchange Online.
- **Sanitized examples.** Example mailboxes throughout use fictional addresses (`user@contoso.com`, `user@domain.com`), and the one-off root `debug_attachment.ps1` is now a parameterized example in `examples/`.
- **Line endings normalized to LF.** The v1 tree showed 31 files as modified with equal insertions and deletions — CRLF vs LF noise. That is gone.

See `CHANGELOG.md` for the full list.

## License

MIT License, copyright 2026 enonymous1. See `LICENSE`.

This is the public export of v2.0.0: the same code as the internal modernization, with a single squashed commit under the `enonymous1` identity and no remote set. The original history is intentionally not included. All examples use fictional addresses, and no tenant data, logs, or exports are included. Before pushing, check `enonymous1/exchange-mailbox-toolkit` on GitHub first: if a repo with that name already exists with content, stop before pushing — this export is meant for an empty repo, not as a replacement for existing history.

---
*Maintained by the automation team.*
