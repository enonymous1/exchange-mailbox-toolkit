# Exchange Online Scripts

PowerShell scripts for managing Exchange Online mailboxes, folders, and messages.

## Scripts

### Ensure_Directory_Structure.ps1
**Purpose:** Export and manage mailbox folder structures with advanced ensure capabilities.

**Features:**
- Export folder hierarchy to CSV or JSON
- Create missing folders from specification files
- Remove unwanted folders (with safe message relocation)
- DryRun mode for safe testing
- Comprehensive audit logging
- Retry logic and error handling

**Basic Usage:**
```powershell
# Export folder structure
.\Ensure_Directory_Structure.ps1 -MailboxEmail "user@domain.com"

# Export as JSON
.\Ensure_Directory_Structure.ps1 -MailboxEmail "user@domain.com" -Format json
```

**Ensure Mode:**
Create a text file with desired folder paths (one per line):
```text
# folders.txt
Inbox
Inbox\Projects
Inbox\Projects\2026
Sent Items
Archive
Archive\2026
```

Then run:
```powershell
# Test what would happen (DryRun)
.\Ensure_Directory_Structure.ps1 -MailboxEmail "user@domain.com" `
    -EnsureFromFile "folders.txt" -DryRun

# Execute the changes
.\Ensure_Directory_Structure.ps1 -MailboxEmail "user@domain.com" `
    -EnsureFromFile "folders.txt" -Force
```

**Parameters:**
- `-MailboxEmail` - Target mailbox email address
- `-OutputDir` - Custom output directory (default: data/output/directory_structures)
- `-Format` - Output format: 'csv' or 'json' (default: csv)
- `-EnsureFromFile` - Path to specification file for folder management
- `-Force` - Skip confirmation prompts
- `-DryRun` - Simulate without making changes
- `-Diagnostics` - Enable detailed diagnostics for troubleshooting
- `-Overwrite` - Overwrite existing output files

**Output:**
- Main export: `data/output/directory_structures/<mailbox>_directory_structure_<timestamp>.csv|json`
- Audit logs (when using ensure mode): `data/output/directory_structures/ensure_logs/`
- Script log: `logs/ensure_directory_structure_<timestamp>.log`

### Export-SharedCalendar.ps1
Export shared calendar events to CSV format.

### Export-Mailbox-Timezones.ps1
Export mailbox timezone configurations.

### Get-Mailboxes.ps1 / Get-GraphMailboxes.ps1
Retrieve and list mailbox information.

### Check_For_Messages.ps1
Search for specific messages based on criteria.

### Move-Messages.ps1
Move messages between folders based on criteria.

### Organize-Mailbox.ps1
Automated mailbox organization and cleanup.

### Export-Mailboxes.ps1
Bulk mailbox export operations.

## Requirements

- PowerShell 7+
- Microsoft Graph PowerShell SDK
  ```powershell
  Install-Module Microsoft.Graph -Scope CurrentUser
  ```
- Exchange Online Management module (for EXO cmdlets)
  ```powershell
  Install-Module ExchangeOnlineManagement -Scope CurrentUser
  ```

## Permissions

### For Export Operations:
- `User.Read.All` - Read user information
- `Mail.Read` - Read mail folders

### For Ensure Operations:
- `User.Read.All` - Read user information
- `Mail.ReadWrite` - Modify mail folders
- `Mail.ReadWrite.Shared` - Modify shared mailboxes (if applicable)

## See Also

- [Utils README](../Utils/README.md) - Documentation for shared utilities
- [Project README](../../README.md) - Overall project documentation

