# Quick Reference - Utility Functions

## GraphHelpers.ps1

### Get-SafeFilename
```powershell
$safe = Get-SafeFilename -Name "user@domain.com"  
# Returns: user_domain_com
```

### Invoke-GraphWithRetry
```powershell
# Automatic retry with exponential backoff
$result = Invoke-GraphWithRetry -Method GET -Uri "users/$userId" -MaxRetries 3
```

### Move-GraphMessage
```powershell
# Move with multiple fallback strategies
$result = Move-GraphMessage `
    -UserId "user@domain.com" `
    -MessageId "AAMk..." `
    -SourceFolderId "inbox" `
    -DestinationFolderId "archive"

if ($result.Success) {
    Write-Host "Moved via: $($result.Strategy)"
} else {
    Write-Warning $result.Error
}
```

### Get-MailFoldersWithPaths
```powershell
# Get all folders with full paths
$folders = Get-MailFoldersWithPaths -UserId "user@domain.com"
foreach ($f in $folders) {
    Write-Host "$($f.Path) - $($f.TotalItemCount) items"
}
```

### Invoke-EnsureFolderPath
```powershell
# Create nested folder path (creates missing parents)
$folderMap = @{}
# ... populate with existing folders ...

$result = Invoke-EnsureFolderPath `
    -UserId "user@domain.com" `
    -FolderPath "Inbox\Projects\2026" `
    -FolderMap $folderMap

if ($result.Success) {
    Write-Host "Created: $($result.Created -join ', ')"
}
```

## LoggingHelpers.ps1

### Set-LogFile
```powershell
# Configure logging destination
Set-LogFile -Path "logs\mylog_$(Get-Date -Format 'yyyyMMdd').log"
```

### Write-LogMessage
```powershell
# Log with different levels
Write-LogMessage "Operation started" -Level 'INFO'
Write-LogMessage "Task completed successfully" -Level 'SUCCESS'
Write-LogMessage "Resource running low" -Level 'WARN'
Write-LogMessage "Operation failed" -Level 'ERROR'
Write-LogMessage "Detailed trace info" -Level 'DEBUG'

# File-only logging (no console)
Write-LogMessage "Sensitive data: $secret" -Level 'DEBUG' -NoConsole
```

### Write-ProgressUpdate
```powershell
# Progress bar with automatic logging
for ($i = 0; $i -lt $total; $i++) {
    $percent = [Math]::Round(($i / $total) * 100)
    Write-ProgressUpdate -Activity "Processing items" `
        -Status "Item $($i+1) of $total" `
        -PercentComplete $percent -Id 1
    # ... do work ...
}
Write-ProgressUpdate -Activity "Processing items" -Completed -Id 1
```

## Initialize-Environment.ps1

### Connect-ExchangeEnvironment
```powershell
# Connect with Graph scopes
$status = Connect-ExchangeEnvironment -GraphScopes @('User.Read.All', 'Mail.Read')

# Connect with Graph and Exchange
$status = Connect-ExchangeEnvironment `
    -GraphScopes @('User.Read.All', 'Mail.ReadWrite') `
    -LoadExchange

# Force re-authentication
$status = Connect-ExchangeEnvironment -GraphScopes @('Mail.Read') -ForceReauth

# Check connection status
if ($status.Graph.Connected) {
    Write-Host "✓ Connected as: $($status.Graph.Account)"
} else {
    Write-Warning "Failed: $($status.Graph.Message)"
}

# Check for errors
if ($status.Errors.Count -gt 0) {
    $status.Errors | ForEach-Object { Write-Warning $_ }
}
```

## Common Patterns

### Standard Script Header
```powershell
#requires -Version 7.0

# Import utilities
$utilsPath = Join-Path -Path $PSScriptRoot -ChildPath '..\Utils'
. (Join-Path $utilsPath 'LoggingHelpers.ps1')
. (Join-Path $utilsPath 'GraphHelpers.ps1')
. (Join-Path $utilsPath 'Initialize-Environment.ps1')

# Configure logging
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$logPath = "logs\myscript_$timestamp.log"
Set-LogFile -Path $logPath

Write-LogMessage "Script started" -Level 'INFO'
```

### Connect to Graph
```powershell
# Connect with error handling
$envStatus = Connect-ExchangeEnvironment -GraphScopes @('Mail.Read') -ErrorAction Stop

if (-not $envStatus.Graph.Connected) {
    Write-LogMessage "Failed to connect: $($envStatus.Graph.Message)" -Level 'ERROR'
    throw "Cannot proceed without Graph connection"
}

Write-LogMessage "Connected as: $($envStatus.Graph.Account)" -Level 'SUCCESS'
```

### Process Items with Progress
```powershell
$items = @(1..100)
$processed = 0

foreach ($item in $items) {
    $processed++
    $percent = [Math]::Round(($processed / $items.Count) * 100)
    
    Write-ProgressUpdate -Activity "Processing items" `
        -Status "Item $processed of $($items.Count)" `
        -PercentComplete $percent -Id 1
    
    try {
        # Process item
        Write-LogMessage "Processed item: $item" -Level 'DEBUG'
    }
    catch {
        Write-LogMessage "Failed to process item $item : $($_.Exception.Message)" -Level 'ERROR'
    }
}

Write-ProgressUpdate -Activity "Processing items" -Completed -Id 1
Write-LogMessage "Processing complete: $processed items" -Level 'SUCCESS'
```

### Safe Graph API Calls
```powershell
# Always use retry wrapper for Graph calls
try {
    $result = Invoke-GraphWithRetry -Method GET -Uri "users/$userId"
    Write-LogMessage "Retrieved user data" -Level 'SUCCESS'
}
catch {
    Write-LogMessage "Failed after retries: $($_.Exception.Message)" -Level 'ERROR'
    throw
}
```

### Create Folder Structure
```powershell
# Build folder map
$folders = Get-MailFoldersWithPaths -UserId $userId
$folderMap = @{}
foreach ($f in $folders) {
    $folderMap[$f.Path.ToLower()] = $f
}

# Ensure folder exists
$desiredPath = "Inbox\Projects\2026"
$result = Invoke-EnsureFolderPath `
    -UserId $userId `
    -FolderPath $desiredPath `
    -FolderMap $folderMap

if ($result.Success -and $result.Created.Count -gt 0) {
    Write-LogMessage "Created: $($result.Created -join ', ')" -Level 'SUCCESS'
}
```

### Error Handling Pattern
```powershell
try {
    # Main operation
    Write-LogMessage "Starting operation" -Level 'INFO'
    $result = Do-Something
    Write-LogMessage "Operation completed" -Level 'SUCCESS'
}
catch {
    Write-LogMessage "Operation failed: $($_.Exception.Message)" -Level 'ERROR'
    
    # Log to file for later analysis
    $errorDetails = @{
        Message = $_.Exception.Message
        StackTrace = $_.ScriptStackTrace
        Timestamp = (Get-Date).ToString('o')
    }
    $errorDetails | ConvertTo-Json | Set-Content "logs\error_$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
    
    throw  # Re-throw if needed
}
finally {
    # Cleanup
    Write-LogMessage "Cleanup completed" -Level 'DEBUG'
}
```

## Tips & Tricks

### Enable Verbose Output
```powershell
$VerbosePreference = 'Continue'
# Now all Write-Verbose calls will display
```

### Debug Mode
```powershell
$DebugPreference = 'Continue'
# Enable debug-level logging
```

### Measure Performance
```powershell
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

# ... operations ...

$stopwatch.Stop()
Write-LogMessage "Completed in $($stopwatch.Elapsed.TotalSeconds)s" -Level 'INFO'
```

### Batch Operations
```powershell
$batchSize = 50
for ($i = 0; $i -lt $items.Count; $i += $batchSize) {
    $batch = $items[$i..[Math]::Min($i + $batchSize - 1, $items.Count - 1)]
    Write-LogMessage "Processing batch $([Math]::Floor($i/$batchSize) + 1)" -Level 'INFO'
    # Process batch
    Start-Sleep -Seconds 1  # Throttle to avoid API limits
}
```

### Dry Run Pattern
```powershell
param([switch]$DryRun)

if ($DryRun) {
    Write-Host "(DryRun) Would perform action" -ForegroundColor Yellow
} else {
    # Actual action
    Write-LogMessage "Performing action" -Level 'INFO'
}
```
