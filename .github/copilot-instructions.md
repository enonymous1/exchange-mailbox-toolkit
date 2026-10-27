# AI Coding Agent Instructions for `exchange-mailbox-toolkit` (v2.0.0)

## Overview
PowerShell 7.2+ tools for Exchange Online / Microsoft 365 mailbox administration, plus a filesystem utility. The shared helpers are also a proper module now: `ExchangeMailboxToolkit/ExchangeMailboxToolkit.psd1` + `.psm1` (13 exported functions).

### Key Directories
- `ExchangeMailboxToolkit/`: the module (manifest + loader).
- `src/`: categorized PowerShell scripts.
  - `Exchange/`: Exchange/M365 operations.
  - `Filesystem/`: local file operations.
  - `Utils/`: shared helpers — Graph, logging, and `Connect-ExchangeEnvironment`.
- `examples/`: diagnostics only (e.g. `Debug-Attachment.ps1`).
- `tests/`: Pester suite (`ExchangeMailboxToolkit.Tests.ps1`).
- `data/`: input/output data files (ignored by git).
- `logs/`: runtime logs (ignored by git).
- `docs/`: documentation and usage notes.

## Developer Workflows
### Running Scripts
1. Open PowerShell in the repository root.
2. Execute scripts using the `pwsh` command. Example:
   ```pwsh
   pwsh -File src/Exchange/Export-SharedCalendar.ps1 -SharedMailboxEmail user@contoso.com
   ```

### Setting Up Environment
Prefer the module. Importing it defines functions only; it never signs in on its own.

```pwsh
Import-Module ./ExchangeMailboxToolkit
$status = Connect-ExchangeEnvironment -GraphScopes 'User.Read.All','Calendars.Read' -LoadExchange -Verbose
$status | Format-List
```

The standalone equivalent is dot-sourcing `src/Utils/Connect-ExchangeEnvironment.ps1` (a loader for the function in `Initialize-Environment.ps1`).

### Logging and Data
- **Logs**: Stored in `logs/`. Ensure sensitive data is excluded.
- **Data**: Input/output files are stored in `data/`.

## Project-Specific Conventions
- **PowerShell Version**: 7.2+ (enforced by `#requires` on scripts and `PowerShellVersion` in the manifest).
- **Paths**: use forward slashes / `Join-Path` segments for filesystem paths, so scripts run under `pwsh` on Windows, macOS, and Linux. Mailbox folder paths use backslashes internally — that is the mailbox hierarchy separator, not a filesystem path; do not "fix" it.
- **Destructive operations**: `Ensure_Directory_Structure.ps1`, `Remove-MailboxFolders.ps1`, and `Import-MoveMessages.ps1` declare `SupportsShouldProcess`. Gate any new destructive call with `ShouldProcess` and keep a `-DryRun` path where one exists. Never add a new `Read-Host` gate to a script that is meant to run non-interactively; `Sort-Inbox.ps1` and the `Export-SharedCalendar.ps1` browser are the deliberate interactive exceptions.
- **No hard-coded tenant data**: examples use `user@contoso.com`. Never reintroduce a real mailbox address.
- **File Organization**: place scripts in their respective folders (`Exchange`, `Filesystem`, etc.); shared functions go in `Utils/` and must also be exported by the module loader and manifest.
- **Verification**: every script must parse cleanly (`[System.Management.Automation.Language.Parser]::ParseFile`), the manifest must pass `Test-ModuleManifest`, and new behaviour gets a Pester case in `tests/`. Live Graph/Exchange behaviour is verified against a real tenant, one mailbox first — never claimed from docs alone.

## Integration Points
- **Exchange Online**: scripts rely on the ExchangeOnlineManagement module (3.0.0+).
- **Microsoft Graph**: Microsoft.Graph module (2.0.0+), via `Connect-ExchangeEnvironment`. The modules are runtime prerequisites checked by `Connect-ExchangeEnvironment`, deliberately not `RequiredModules` in the manifest, so the module still imports without them.

## Examples
### Exporting Shared Calendar
```pwsh
pwsh -File src/Exchange/Export-SharedCalendar.ps1 -SharedMailboxEmail user@contoso.com
```

### Generating File List
```pwsh
pwsh -File src/Filesystem/FileList_Generator.ps1 -Path ./src -OutputFile data/output/file_list.csv
```

---

For more details, refer to the [README.md](../README.md) files in the respective directories.
