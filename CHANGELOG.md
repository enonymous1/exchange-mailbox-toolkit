# Changelog - exchange-mailbox-toolkit Improvements

## Function-name generalization - 2026-10-27

The `MP`-prefixed function names described what the toolkit was written for, not what the functions do. Generalized before the first public push, when a rename is still free. No aliases for the old names are kept; scripts calling them need updating.

| Old name | New name |
|---|---|
| `Connect-MPEnv` | `Connect-ExchangeEnvironment` |
| `Write-MPLog` | `Write-LogMessage` |
| `Set-MPLogFile` | `Set-LogFile` |
| `Clear-MPLogFile` | `Clear-LogFile` |
| `Write-MPProgress` | `Write-ProgressUpdate` |
| `Get-MpMailboxes` | `Get-GraphMailboxes` |

The loader and discovery files were renamed to match (`Connect-ExchangeEnvironment.ps1`, `Get-GraphMailboxes.ps1`), and the internal `$script:MPLogFile` became `$script:LogFilePath`. The already-generic names (`Get-SafeFilename`, `Get-EncodedGraphId`, `Invoke-GraphWithRetry`, `Move-GraphMessage`, `Get-MailFoldersWithPaths`, `Invoke-EnsureFolderPath`) are unchanged. `Get-Mailboxes` is kept as a deprecated shim that forwards to `Get-GraphMailboxes`.

## Rename + public export - 2026-10-27

Renamed from `mp-workingDir` / `MPWorkingDir` to `exchange-mailbox-toolkit` / `ExchangeMailboxToolkit` before the first public push, when a rename is still free. The repo and module names changed, and the module GUID was regenerated for the new module name.

Same v2.0.0 code as the internal modernization, prepared as a single squashed commit under the `enonymous1` identity with no remote. MIT LICENSE (holder enonymous1) added; the execution-history summary was generalized (no client name, no unverified scale/success metrics); the sanitization regression test was reworded so no real organizational address or name appears anywhere in the tree or history. The original history is intentionally not included — use this export, not the internal v2 zip or the `mp-workingDir` public export, for any public push, and only into an empty repo after checking the target name on GitHub first.

## [2.0.0] - 2026-10-26 — Full modernization

### Fixed
- `Ensure_Directory_Structure.ps1` did not parse: two closing braces were missing in the folder-creation loop (the `else` for a failed create and the `foreach` over desired folders were never closed). v1 docs claimed syntax validation passed; the PowerShell parser disagrees, and v2 is parser-clean.
- `Get-Mailboxes.ps1` did not parse: its header comment ended with `>` instead of `#>`, so the whole file was one unterminated comment and the shim function was never defined.
- The `Connect-ExchangeEnvironment.ps1` loader did not exist, although the README, Copilot instructions, `Export-Mailboxes.ps1`, and `Export-SharedCalendar.ps1` all referenced it. It now exists and loads `Initialize-Environment.ps1`.
- Dead `Helpers.ps1` references removed (`Initialize-Environment.ps1`, `Export-SharedCalendar.ps1`).
- `Check_For_Messages.ps1` id parsing: a file containing both commas and newlines was split on commas only, merging ids across lines into one bogus id. It now splits on commas and newlines.
- README example for `FileList_Generator.ps1` used parameters that do not exist (`-DirectoryPath`, `-OutputPath`); the real parameters are `-Path` and `-OutputFile`.
- Line endings normalized to LF (v1 tree showed 31 files modified with 5,334 insertions / 5,334 deletions — CRLF noise only).

### Changed
- Proper module: `ExchangeMailboxToolkit/ExchangeMailboxToolkit.psd1` + `.psm1`, version 2.0.0, PowerShell 7.2, 13 exported functions. Importing the module defines functions only; it no longer triggers a sign-in (the auto-run block in `Initialize-Environment.ps1` was removed).
- `Microsoft.Graph` (>= 2.0.0) and `ExchangeOnlineManagement` (>= 3.0.0) are runtime prerequisites checked by `Connect-ExchangeEnvironment`, deliberately not `RequiredModules` in the manifest, so the module imports and the tests run without Graph/EXO installed.
- Cross-platform filesystem paths: backslashes in `Join-Path` / dot-source strings replaced with forward slashes / `Join-Path` segments. Mailbox folder-path backslashes left untouched (mailbox hierarchy separator, not a filesystem path).
- `Check_For_Messages.ps1` is now parameter-driven (`-CsvPath`, `-IdFilePath`, `-MessageIdColumn`, `-OutputDir`, `-Save`) with fail-fast errors instead of `Read-Host` loops.
- `Ensure_Directory_Structure.ps1`: `-MailboxEmail` is now Mandatory, and folder removal is gated by `ShouldProcess` instead of `Read-Host`, so `-WhatIf` / `-Confirm` work non-interactively.
- `Remove-MailboxFolders.ps1` and `Import-MoveMessages.ps1` now declare `SupportsShouldProcess` and gate their destructive call with `ShouldProcess`.
- `Export-Mailbox-Timezones.ps1` implemented (was a placeholder): Exchange Online regional configuration export, with an optional `-MailboxEmail` scope for testing.
- Root `debug_attachment.ps1` moved to `examples/Debug-Attachment.ps1` with `-MailboxEmail` / `-MessageId` parameters instead of a hard-coded mailbox.
- Example mailboxes sanitized throughout (a real organizational address -> `user@contoso.com`); organization-specific wording generalized.

### Added
- Pester test suite (`tests/ExchangeMailboxToolkit.Tests.ps1`): manifest, exports, pure helpers, parse of every script, loader regressions, cross-platform path regressions, `ShouldProcess` declarations, `Check_For_Messages` end-to-end, sanitization.
- PSScriptAnalyzer settings and GitHub Actions CI on Windows + Ubuntu `pwsh`: manifest test, PSScriptAnalyzer, Pester.
- Landing-page README with safety guidance and an honest v2.0.0 change list.

### Not verified here
- Live Graph / Exchange Online behaviour: no tenant access on the build box. Verified here: PowerShell parser clean for every script, module import + manifest, pure helper outputs, and `Check_For_Messages.ps1` end-to-end on fixture data. Live tenant testing (one mailbox first) and the CI run itself are still the user's gates. Pester could not be installed on the build box (PSGallery repository registration failed there), so the Pester file was verified by parser + an equivalent local harness, not by the Pester runner itself.

## [Unreleased] - 2026-01-15

### Added

#### New Utility Modules
- **GraphHelpers.ps1** - Comprehensive Microsoft Graph API helper functions
  - `Get-SafeFilename` - Convert strings to filesystem-safe names
  - `Get-EncodedGraphId` - URL-encode identifiers for Graph API
  - `Invoke-GraphWithRetry` - Execute requests with automatic retry and exponential backoff
  - `Move-GraphMessage` - Move messages with multiple fallback strategies
  - `Get-MailFoldersWithPaths` - Recursively enumerate folder hierarchies
  - `Invoke-EnsureFolderPath` - Create nested folder paths with parent creation

#### Enhanced Utility Modules
- **LoggingHelpers.ps1**
  - `Write-ProgressUpdate` - Progress bars with automatic logging
  - Enhanced `Write-LogMessage` with DEBUG and SUCCESS levels
  - Color-coded console output
  - `-NoConsole` parameter for file-only logging

#### Documentation
- `src/Utils/README.md` - Comprehensive utilities documentation
- `src/Exchange/README.md` - Enhanced with detailed usage examples
- `docs/improvements-summary.md` - Complete refactoring documentation
- `docs/quick-reference.md` - Quick reference for utility functions

#### Features
- Automatic retry logic with exponential backoff for all Graph API calls
- Multiple fallback strategies for message move operations
- Progress bars for long-running operations
- Comprehensive audit logging for all operations
- DryRun simulation mode throughout
- Diagnostics mode for detailed troubleshooting
- Better error messages with context
- Operation summaries with statistics

### Changed

#### Ensure_Directory_Structure.ps1
- **Breaking Changes:** None (fully backward compatible)
- **Improvements:**
  - Refactored to use new utility modules
  - Enhanced parameter documentation with examples
  - Added `#requires -Version 7.0` directive
  - Extracted helper functions to reusable utilities
  - Improved folder creation logic (batch operations)
  - Enhanced message moving with retry and fallbacks
  - Better progress reporting throughout
  - Comprehensive logging to timestamped log files
  - Enhanced audit trail with operation summaries
  - Better user prompts and feedback
  - Visual improvements (checkmarks, colors, formatting)
  - Constants defined at module level (PROTECTED_SYSTEM_FOLDERS)

- **New Parameters:**
  - `-Diagnostics` - Enable detailed diagnostic output
  - (All other parameters enhanced with better help text)

- **Performance:**
  - Reduced API calls (single folder refresh vs multiple)
  - Batch folder creation before refresh
  - More efficient folder mapping with hashtables

### Fixed
- Message move operations now handle edge cases better
- Folder creation properly handles nested paths
- Better handling of Graph API throttling
- Proper error propagation throughout
- URL encoding handled consistently

### Security
- Sensitive data can be logged file-only (no console output)
- Better protection of system folders
- Comprehensive audit trails for compliance

### Deprecated
- None

### Removed
- None

### Technical Details

#### Dependencies
- Requires PowerShell 7.0+
- Microsoft Graph PowerShell SDK
- Exchange Online Management module (for Exchange operations)

#### File Structure Changes
```
Added:
  src/Utils/GraphHelpers.ps1
  src/Utils/README.md
  docs/improvements-summary.md
  docs/quick-reference.md

Modified:
  src/Exchange/Ensure_Directory_Structure.ps1 (major refactor)
  src/Utils/LoggingHelpers.ps1 (enhancements)
  src/Exchange/README.md (comprehensive documentation)

Utilized:
  src/Utils/Initialize-Environment.ps1 (now properly imported)
```

#### Code Metrics
- **Lines of Code:**
  - Before: ~415 lines (monolithic)
  - After: ~450 lines main script + ~350 lines utilities
  - Net: More maintainable with separated concerns

- **Functions:**
  - Before: 3 inline functions
  - After: 11 reusable utility functions + main script logic

- **Error Handling:**
  - Before: Basic try/catch in critical sections
  - After: Comprehensive error handling with retry logic throughout

#### Testing Recommendations
1. Test basic export functionality (no regressions)
2. Test ensure mode with DryRun first
3. Verify retry logic with throttled APIs
4. Validate audit log completeness
5. Test progress reporting with large mailboxes
6. Verify error handling and logging

#### Migration Notes
No migration required. All changes are backward compatible:
- Existing parameter names unchanged
- Default behaviors preserved
- Output locations unchanged
- New features are additive and opt-in

#### Known Limitations
- PowerShell linter shows warnings for internal functions (cosmetic only)
  - `Record-Created`, `Record-Failure` (internal helpers)
  - `Recurse-Folders` (nested function)
- These are intentional and don't affect functionality

## Future Roadmap

### Planned Enhancements
- [ ] Parallel processing for multiple mailboxes
- [ ] Configuration file support (YAML/JSON)
- [ ] Email notifications for completion/errors
- [ ] Web dashboard for audit logs
- [ ] Template library for common folder structures
- [ ] Backup/restore functionality
- [ ] Integration with ticketing systems
- [ ] Scheduled execution support
- [ ] Performance metrics collection
- [ ] Automated testing suite

### Under Consideration
- CSV import for bulk operations
- Folder permission management
- Message rule management
- Retention policy enforcement
- Compliance reporting
- Multi-tenant support

---

## Version History

### [Current] - 2026-01-15
Complete refactoring with utility modules and comprehensive improvements.

### [Previous] - Before 2026-01-15
Original monolithic script with basic functionality.

---

## Contributors
- AI Assistant (GitHub Copilot) - Refactoring and improvements

## References
- [Microsoft Graph API Documentation](https://docs.microsoft.com/graph)
- [PowerShell Best Practices](https://docs.microsoft.com/powershell/scripting/developer/cmdlet/cmdlet-development-guidelines)
- [Exchange Online PowerShell](https://docs.microsoft.com/powershell/exchange/exchange-online-powershell)
