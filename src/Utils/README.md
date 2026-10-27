# Utilities Module

This directory contains reusable PowerShell utilities for Exchange/M365 operations and general automation tasks.

## Available Modules

### LoggingHelpers.ps1
Provides consistent logging functionality across all scripts.

**Functions:**
- `Set-LogFile` - Configure log file path
- `Clear-LogFile` - Clear log file configuration
- `Write-LogMessage` - Write timestamped log entries with levels (INFO, WARN, ERROR, DEBUG, SUCCESS)
- `Write-ProgressUpdate` - Write progress bars with automatic logging

**Example:**
```powershell
# Import and configure logging
. "$PSScriptRoot\Utils\LoggingHelpers.ps1"
Set-LogFile -Path "logs\mylog.log"

# Write log messages
Write-LogMessage -Message "Operation started" -Level 'INFO'
Write-LogMessage -Message "Task completed successfully" -Level 'SUCCESS'
Write-LogMessage -Message "Warning: Resource low" -Level 'WARN'
```

### GraphHelpers.ps1
Reusable functions for Microsoft Graph API operations with retry logic and error handling.

**Functions:**
- `Get-SafeFilename` - Convert strings to safe filenames
- `Get-EncodedGraphId` - URL-encode Graph API identifiers
- `Invoke-GraphWithRetry` - Execute Graph requests with exponential backoff
- `Move-GraphMessage` - Move messages with multiple fallback strategies
- `Get-MailFoldersWithPaths` - Recursively enumerate mail folder hierarchy
- `Invoke-EnsureFolderPath` - Create folder paths with parent hierarchy

**Features:**
- Automatic retry with exponential backoff for throttling (429) and transient errors (5xx)
- Multiple fallback strategies for problematic operations
- Progress reporting for long-running operations
- DryRun support for safe testing

**Example:**
```powershell
# Import helper
. "$PSScriptRoot\Utils\GraphHelpers.ps1"

# Get all folders for a mailbox
$folders = Get-MailFoldersWithPaths -UserId "user@domain.com"

# Move a message with automatic retry
$result = Move-GraphMessage -UserId "user@domain.com" `
    -MessageId "AAMk..." `
    -SourceFolderId "inbox" `
    -DestinationFolderId "folder123"

if ($result.Success) {
    Write-Host "Moved via strategy: $($result.Strategy)"
}
```

### Initialize-Environment.ps1
Centralized authentication and environment setup for Microsoft Graph and Exchange Online.

**Function:**
- `Connect-ExchangeEnvironment` - Establish connections with session checking

**Parameters:**
- `-GraphScopes` - Array of Graph API permissions
- `-LoadExchange` - Switch to connect to Exchange Online
- `-ForceReauth` - Force re-authentication
- `-LogFile` - Optional log file path

**Returns:**
Status object with connection details and any errors/warnings.

**Example:**
```powershell
# Import and connect
. "$PSScriptRoot\Utils\Initialize-Environment.ps1"

# Connect with specific scopes
$status = Connect-ExchangeEnvironment -GraphScopes @('User.Read.All', 'Mail.ReadWrite') -LoadExchange

if ($status.Graph.Connected) {
    Write-Host "Connected as: $($status.Graph.Account)"
}

# Check for errors
if ($status.Errors.Count -gt 0) {
    foreach ($err in $status.Errors) {
        Write-Warning $err
    }
}
```

## Best Practices

### Using Multiple Utilities
Scripts should import all needed utilities at the start:

```powershell
#requires -Version 7.0

# Import utilities
$utilsPath = Join-Path -Path $PSScriptRoot -ChildPath '..\Utils'
. (Join-Path $utilsPath 'LoggingHelpers.ps1')
. (Join-Path $utilsPath 'GraphHelpers.ps1')
. (Join-Path $utilsPath 'Initialize-Environment.ps1')

# Configure logging
$logPath = "logs\myscript_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
Set-LogFile -Path $logPath

# Connect to services
$envStatus = Connect-ExchangeEnvironment -GraphScopes @('Mail.Read') -ErrorAction Stop
```

### Error Handling
Always check connection status before proceeding:

```powershell
if (-not $envStatus.Graph.Connected) {
    Write-LogMessage -Message "Failed to connect: $($envStatus.Graph.Message)" -Level 'ERROR'
    return
}
```

### Progress Reporting
Use `Write-ProgressUpdate` for long operations:

```powershell
$total = $items.Count
for ($i = 0; $i -lt $total; $i++) {
    $percent = [Math]::Round(($i / $total) * 100)
    Write-ProgressUpdate -Activity "Processing items" `
        -Status "Item $($i+1) of $total" `
        -PercentComplete $percent -Id 1
    
    # Process item
}
Write-ProgressUpdate -Activity "Processing items" -Completed -Id 1
```

### Retry Logic
Use `Invoke-GraphWithRetry` for all Graph API calls to handle throttling:

```powershell
# Instead of:
$result = Invoke-MgGraphRequest -Method GET -Uri "users/$userId"

# Use:
$result = Invoke-GraphWithRetry -Method GET -Uri "users/$userId"
```

## Requirements

- PowerShell 7+
- Microsoft Graph PowerShell SDK (`Install-Module Microsoft.Graph`)
- Exchange Online Management module (for Exchange operations)

## Testing

All utilities support verbose output for debugging:

```powershell
$VerbosePreference = 'Continue'
Connect-ExchangeEnvironment -GraphScopes @('Mail.Read') -Verbose
```

## Contributing

When adding new utility functions:
1. Include proper comment-based help with `.SYNOPSIS`, `.DESCRIPTION`, `.PARAMETER`, `.EXAMPLE`
2. Use `[CmdletBinding()]` for advanced function features
3. Export functions explicitly in the module export section
4. Add comprehensive error handling with meaningful messages
5. Update this README with documentation
