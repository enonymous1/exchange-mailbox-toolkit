<#
.SYNOPSIS
Export and manage mailbox folder (directory) structures.

.DESCRIPTION
This script provides comprehensive folder structure management for Exchange Online mailboxes:
- Exports folder hierarchy to CSV or JSON format
- Creates missing folders from a specification file
- Removes folders not in the specification (with message relocation)
- Supports dry-run mode for safe testing

The script connects to Microsoft Graph and recursively enumerates mail folders, building full paths
for each folder. Results are exported to `data/output/directory_structures` with timestamped filenames.

.PARAMETER MailboxEmail
Mailbox email address to inspect. If not provided, user will be prompted.

.PARAMETER OutputDir
Optional custom output directory. If omitted, uses repo `data/output/directory_structures`.

.PARAMETER Format
Output format: 'csv' (default) or 'json'.

.PARAMETER EnsureFromFile
Path to a text file containing desired folder paths (one per line). When provided, the script will:
- Create any missing folders
- Remove folders not in the list (after moving messages to Inbox)
Lines starting with # are treated as comments.

.PARAMETER Force
Skip confirmation prompts for destructive operations (folder deletion).

.PARAMETER DryRun
Simulate operations without making actual changes. Useful for testing.

.PARAMETER Diagnostics
Enable detailed diagnostics for failed message move operations.

.PARAMETER Overwrite
If specified and an output file already exists with the same name, it will be overwritten.

.EXAMPLE
# Export folder structure to CSV
.\Ensure_Directory_Structure.ps1 -MailboxEmail "user@domain.com"

.EXAMPLE
# Ensure folders match specification file (dry-run)
.\Ensure_Directory_Structure.ps1 -MailboxEmail "user@domain.com" -EnsureFromFile "folders.txt" -DryRun

.EXAMPLE
# Enforce folder structure without prompts
.\Ensure_Directory_Structure.ps1 -MailboxEmail "user@domain.com" -EnsureFromFile "folders.txt" -Force

.NOTES
Requires Microsoft Graph PowerShell SDK with Mail.Read (or Mail.ReadWrite for ensure operations).
System folders (Inbox, Sent Items, etc.) are protected from deletion.
#>
#requires -Version 7.2

[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [Parameter(Position=0, Mandatory=$true, HelpMessage='Mailbox email to inspect')]
    [string]$MailboxEmail,

    [Parameter(HelpMessage='Output directory (optional)')]
    [string]$OutputDir,

    [Parameter(HelpMessage='Output format')]
    [ValidateSet('csv','json')]
    [string]$Format = 'csv',

    [Parameter(HelpMessage='Text file containing desired folder paths (one per line)')]
    [string]$EnsureFromFile,

    [Parameter(HelpMessage='Skip confirmation prompts')]
    [switch]$Force,

    [Parameter(HelpMessage='Simulate operations without making changes')]
    [switch]$DryRun,

    [Parameter(HelpMessage='Enable detailed diagnostics for troubleshooting')]
    [switch]$Diagnostics,

    [Parameter(HelpMessage='Overwrite existing output files')]
    [switch]$Overwrite
)



# Constants
$PROTECTED_SYSTEM_FOLDERS = @(
    'Inbox', 'Drafts', 'Sent Items', 'Deleted Items', 
    'Junk Email', 'Outbox', 'Archive', 'Conversation History'
)

# Import utilities
$utilsPath = Join-Path -Path $PSScriptRoot -ChildPath '../Utils'
try {
    . (Join-Path $utilsPath 'LoggingHelpers.ps1')
    . (Join-Path $utilsPath 'GraphHelpers.ps1')
    . (Join-Path $utilsPath 'Initialize-Environment.ps1')
}
catch {
    Write-Warning "Failed to import utilities: $($_.Exception.Message). Some features may not work."
}

# Initialize logging
$repoRoot = Resolve-Path -Path (Join-Path $PSScriptRoot '../..')
$logDir = Join-Path $repoRoot 'logs'
if (!(Test-Path $logDir)) { 
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null 
}
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$logPath = Join-Path $logDir "ensure_directory_structure_$timestamp.log"
Set-LogFile -Path $logPath

Write-LogMessage -Message "=== Ensure_Directory_Structure Script Started ===" -Level 'INFO'
Write-Host "`nMailbox Folder Structure Manager" -ForegroundColor Cyan
Write-Host "================================" -ForegroundColor Cyan

# v2.0.0: -MailboxEmail is now Mandatory. PowerShell prompts for it in an
# interactive session and fails fast in a non-interactive one, instead of
# the script asking with Read-Host and then continuing on a blank answer.

Write-LogMessage -Message "Target mailbox: $MailboxEmail" -Level 'INFO'

# Ensure Graph connection using utility
$requiredScopes = @('User.Read.All', 'Mail.Read')
if ($EnsureFromFile -or $Force) {
    $requiredScopes += 'Mail.ReadWrite'
    $requiredScopes += 'Mail.ReadWrite.Shared'
}

Write-LogMessage -Message "Checking Microsoft Graph connection..." -Level 'INFO'
$envStatus = Connect-ExchangeEnvironment -GraphScopes $requiredScopes -ErrorAction Stop

if (-not $envStatus.Graph.Connected) {
    Write-LogMessage -Message "Failed to connect to Microsoft Graph: $($envStatus.Graph.Message)" -Level 'ERROR'
    Write-Error "Failed to connect to Microsoft Graph. Exiting."
    return
}

Write-LogMessage -Message "Connected to Graph as: $($envStatus.Graph.Account)" -Level 'SUCCESS'

# Look up user
Write-LogMessage -Message "Looking up user: $MailboxEmail" -Level 'INFO'
$TargetUser = Get-MgUser -Filter "mail eq '$MailboxEmail'" -ConsistencyLevel eventual -ErrorAction SilentlyContinue | Select-Object -First 1

if (-not $TargetUser) {
    Write-LogMessage -Message "Primary mail lookup failed. Checking proxy addresses..." -Level 'WARN'
    $TargetUser = Get-MgUser -Filter "proxyAddresses/any(p:p eq 'smtp:$MailboxEmail')" -ConsistencyLevel eventual -ErrorAction SilentlyContinue | Select-Object -First 1
}

if (-not $TargetUser) { 
    Write-LogMessage -Message "Could not find user for $MailboxEmail" -Level 'ERROR'
    Write-Error "Could not find user for $MailboxEmail"
    return 
}

Write-LogMessage -Message "Found user: $($TargetUser.DisplayName) (ID: $($TargetUser.Id))" -Level 'SUCCESS'

# Enumerate folders using helper function
Write-LogMessage -Message "Enumerating mail folders..." -Level 'INFO'
$folders = Get-MailFoldersWithPaths -UserId $TargetUser.Id -ProgressActivity "Scanning $MailboxEmail"
Write-LogMessage -Message "Found $($folders.Count) folders" -Level 'SUCCESS'

# If an ensure file is provided, process create/delete operations
if ($EnsureFromFile) {
    Write-LogMessage -Message "Ensure mode enabled with file: $EnsureFromFile" -Level 'INFO'
    
    if (!(Test-Path $EnsureFromFile)) { 
        Write-LogMessage -Message "Ensure file not found: $EnsureFromFile" -Level 'ERROR'
        Write-Error "Ensure file not found: $EnsureFromFile"
        return 
    }

    # Load desired folder paths
    $desired = Get-Content -Path $EnsureFromFile | ForEach-Object { $_.Trim() } | Where-Object { $_ -and -not $_.StartsWith('#') }
    if ($desired.Count -eq 0) { 
        Write-LogMessage -Message "No folder paths found in ensure file." -Level 'ERROR'
        Write-Error "No folder paths found in ensure file."
        return 
    }

    # Normalize paths
    $desiredNormalized = $desired | ForEach-Object { ($_ -replace '/','\\').Trim() }
    Write-LogMessage -Message "Loaded $($desiredNormalized.Count) desired folders from specification" -Level 'INFO'

    # Build current folder map
    $folderByPath = @{}
    foreach ($f in $folders) { 
        $folderByPath[$f.Path.ToLower()] = $f 
    }

    # Initialize audit tracking
    $auditCreated = [System.Collections.ArrayList]::new()
    $auditMoved = [System.Collections.ArrayList]::new()
    $auditDeleted = [System.Collections.ArrayList]::new()
    $auditFailed = [System.Collections.ArrayList]::new()

    # Helper functions for audit recording
    function Add-CreatedAudit { 
        param($path, $id)
        [void]$auditCreated.Add([PSCustomObject]@{ 
            FolderPath = $path
            FolderId = $id
            Timestamp = (Get-Date).ToString('o')
        })
        Write-LogMessage -Message "Created folder: $path" -Level 'SUCCESS'
    }
    
    function Add-FailureAudit { 
        param($type, $id, $msg, $tried)
        [void]$auditFailed.Add([PSCustomObject]@{ 
            Type = $type
            Id = $id
            Error = $msg
            Tried = $tried
            Timestamp = (Get-Date).ToString('o')
        })
        Write-LogMessage -Message "Failed $type : $id - $msg" -Level 'ERROR'
    }

    # Diagnostics helper for troubleshooting message move failures
    function Get-MessageDiagnostics {
        param(
            [string]$mid,
            [string]$folderId
        )
        
        Write-LogMessage -Message "Running diagnostics for message: $mid" -Level 'DEBUG'
        $encMid = Get-EncodedGraphId -Value $mid
        $tries = [System.Collections.ArrayList]::new()

        # Try direct message GET
        $uri1 = "users/$($TargetUser.Id)/messages/$encMid"
        try { 
            $res = Invoke-MgGraphRequest -Method GET -Uri $uri1 -ErrorAction Stop
            [void]$tries.Add([PSCustomObject]@{ 
                Uri = $uri1
                Status = 'OK'
                Detail = ($res | ConvertTo-Json -Depth 3 -Compress)
            })
        } 
        catch { 
            [void]$tries.Add([PSCustomObject]@{ 
                Uri = $uri1
                Status = 'Error'
                Detail = $_.Exception.Message
            })
        }

        # Try folder-scoped GET
        $uri2 = "users/$($TargetUser.Id)/mailFolders/$folderId/messages/$encMid"
        try { 
            $res2 = Invoke-MgGraphRequest -Method GET -Uri $uri2 -ErrorAction Stop
            [void]$tries.Add([PSCustomObject]@{ 
                Uri = $uri2
                Status = 'OK'
                Detail = ($res2 | ConvertTo-Json -Depth 3 -Compress)
            })
        } 
        catch { 
            [void]$tries.Add([PSCustomObject]@{ 
                Uri = $uri2
                Status = 'Error'
                Detail = $_.Exception.Message
            })
        }

        # List messages in folder
        try { 
            $list = Get-MgUserMailFolderMessage -UserId $TargetUser.Id -MailFolderId $folderId `
                -Property 'Id,Subject' -PageSize 50 -All -ErrorAction SilentlyContinue
            $found = $list | Where-Object { $_.Id -eq $mid }
            [void]$tries.Add([PSCustomObject]@{ 
                Uri = 'ListFolderMessages'
                Status = if ($found) { 'Found' } else { 'NotFound' }
                Detail = if ($found) { ($found | Select-Object Id, Subject | ConvertTo-Json -Depth 3 -Compress) } else { '' }
            })
        } 
        catch { 
            [void]$tries.Add([PSCustomObject]@{ 
                Uri = 'ListFolderMessages'
                Status = 'Error'
                Detail = $_.Exception.Message
            })
        }

        return $tries
    }

    # Ensure folder path exists (create missing parents)
    Write-LogMessage -Message "Ensuring desired folders exist..." -Level 'INFO'
    $createCount = 0
    foreach ($d in $desiredNormalized) {
        $result = Invoke-EnsureFolderPath -UserId $TargetUser.Id -FolderPath $d `
            -FolderMap $folderByPath -DryRun:$DryRun
        
        if ($result.Success) {
            foreach ($created in $result.Created) {
                Add-CreatedAudit -path $created -id $folderByPath[$created.ToLower()].Id
                $createCount++
            }
        }
        else {
            Add-FailureAudit -type 'CreateFolder' -id $d -msg $result.Error -tried ''
        }
    }
    
    if ($createCount -gt 0) {
        Write-LogMessage -Message "Created $createCount folder(s)" -Level 'SUCCESS'
    }
    else {
        Write-LogMessage -Message "All desired folders already exist" -Level 'INFO'
    }

    # Refresh folder list after creations
    if ($createCount -gt 0 -and -not $DryRun) {
        Write-LogMessage -Message "Refreshing folder list..." -Level 'INFO'
        $folders = Get-MailFoldersWithPaths -UserId $TargetUser.Id
        $folderByPath = @{}
        foreach ($f in $folders) { 
            $folderByPath[$f.Path.ToLower()] = $f 
        }
    }

    # Determine deletion candidates (folders not in desired list)
    Write-LogMessage -Message "Identifying folders for removal..." -Level 'INFO'
    $desiredSet = @{}
    foreach ($d in $desiredNormalized) { 
        $desiredSet[$d.ToLower()] = $true 
    }

    $candidates = [System.Collections.ArrayList]::new()
    foreach ($f in $folders) {
        $p = $f.Path
        if ($desiredSet.ContainsKey($p.ToLower())) { 
            continue 
        }
        
        # Never delete top-level system folders
        $top = ($p -split '\\')[0]
        if ($PROTECTED_SYSTEM_FOLDERS -contains $top) { 
            # Allow deleting subfolders under system top-levels
            if ($p -eq $top) { 
                continue 
            }
        }
        [void]$candidates.Add($f)
    }
    
    # Sort candidates: deepest folders first (to delete children before parents)
    # This prevents parent deletion from cascading and causing NotFound errors
    $candidates = $candidates | Sort-Object { ($_.Path -split '\\').Count } -Descending

    if ($candidates.Count -eq 0) { 
        Write-LogMessage -Message "No extra folders to remove" -Level 'INFO'
    }
    else {
        Write-LogMessage -Message "Found $($candidates.Count) folders not in the desired list" -Level 'WARN'
        Write-LogMessage -Message "Folders will be processed deepest-first to avoid cascading deletions" -Level 'INFO'
        foreach ($c in $candidates) { 
            Write-Host "  • $($c.Path) (Items: $($c.TotalItemCount))" -ForegroundColor Yellow
        }

        if (-not $Force -and -not $DryRun) {
            Write-Host "`nThese folders will be DELETED after moving their messages to Inbox." -ForegroundColor Yellow
            # v2.0.0: -WhatIf / -Confirm now work here via ShouldProcess, so
            # this gate also works non-interactively instead of Read-Host.
            if (-not $PSCmdlet.ShouldProcess("$($candidates.Count) folders in $MailboxEmail", 'Remove folders not in specification')) {
                Write-LogMessage -Message "User cancelled folder removal" -Level 'INFO'
                Write-Host "Skipping removal." -ForegroundColor Gray
                $candidates.Clear()
            }
        }
        elseif ($Force) {
            Write-LogMessage -Message "Force mode enabled: will remove listed folders" -Level 'WARN'
        }
        elseif ($DryRun) {
            Write-LogMessage -Message "DryRun mode: simulating folder removal" -Level 'INFO'
        }

        # Find Inbox
        $inboxObj = $folders | Where-Object { $_.Path -ieq 'Inbox' } | Select-Object -First 1
        if (-not $inboxObj -and $candidates.Count -gt 0) { 
            Write-LogMessage -Message "Inbox folder not found; aborting removal" -Level 'ERROR'
            Write-Warning "Inbox folder not found; aborting removal."
            $candidates.Clear()
        }
        $inboxId = $inboxObj.Id

        # Process each candidate for removal
        $processed = 0
        foreach ($c in $candidates) {
            $processed++
            $percentComplete = [Math]::Round(($processed / $candidates.Count) * 100)
            Write-ProgressUpdate -Activity "Processing folder removals" -Status "Folder $processed of $($candidates.Count): $($c.Path)" `
                -PercentComplete $percentComplete -Id 2

            Write-LogMessage -Message "Processing removal of folder: $($c.Path)" -Level 'INFO'
            
            # Query folder fresh by path to get valid current ID
            Write-Verbose "Querying fresh folder ID for: $($c.Path)"
            $freshFolder = Get-FreshFolderByPath -UserId $TargetUser.Id -FolderPath $c.Path
            
            if (-not $freshFolder) {
                Write-LogMessage -Message "Folder not found (may have been deleted): $($c.Path)" -Level 'WARN'
                [void]$auditDeleted.Add([PSCustomObject]@{ 
                    FolderPath = $c.Path
                    FolderId = 'N/A'
                    Status = 'AlreadyDeleted'
                    Timestamp = (Get-Date).ToString('o')
                })
                continue
            }

            $freshId = $freshFolder.id
            Write-Verbose "Fresh folder ID: $freshId (TotalItems: $($freshFolder.totalItemCount))"
            
            # Get messages in folder using fresh ID
            $msgs = @()
            try {
                $msgs = Get-MgUserMailFolderMessage -UserId $TargetUser.Id -MailFolderId $freshId `
                    -PageSize 50 -All -ErrorAction Stop
            } 
            catch { 
                if ($_.Exception.Message -match 'NotFound') {
                    Write-LogMessage -Message "Folder no longer exists (deleted during processing): $($c.Path)" -Level 'WARN'
                    [void]$auditDeleted.Add([PSCustomObject]@{ 
                        FolderPath = $c.Path
                        FolderId = $freshId
                        Status = 'AlreadyDeleted'
                        Timestamp = (Get-Date).ToString('o')
                    })
                    continue
                }
                Write-LogMessage -Message "Could not enumerate messages: $($_.Exception.Message)" -Level 'WARN'
                $msgs = @() 
            }

            if ($msgs.Count -gt 0) {
                Write-LogMessage -Message "Found $($msgs.Count) messages in folder; moving to Inbox" -Level 'INFO'
                $msgProcessed = 0
                
                foreach ($m in $msgs) {
                    $msgProcessed++
                    if ($DryRun) { 
                        Write-Host "  (DryRun) Would move message $($msgProcessed)/$($msgs.Count) to Inbox" -ForegroundColor Gray
                        continue 
                    }

                    try {
                        # Use the improved Move-GraphMessage helper with fresh folder ID
                        $moveResult = Move-GraphMessage -UserId $TargetUser.Id -MessageId $m.Id `
                            -SourceFolderId $freshId -DestinationFolderId $inboxId

                        if ($moveResult.Success) {
                            Write-Verbose "  Moved message $msgProcessed/$($msgs.Count) via $($moveResult.Strategy)"
                            [void]$auditMoved.Add([PSCustomObject]@{ 
                                MessageId = $m.Id
                                FromFolder = $c.Path
                                ToFolder = 'Inbox'
                                Status = 'Moved'
                                Strategy = $moveResult.Strategy
                                Tried = $moveResult.TriedUris
                            })
                        }
                        else {
                            # Check if it's a NotFound error (message already moved or deleted)
                            if ($moveResult.Error -match 'NotFound') {
                                Write-Verbose "  Message already moved or deleted: $($m.Id)"
                                [void]$auditMoved.Add([PSCustomObject]@{ 
                                    MessageId = $m.Id
                                    FromFolder = $c.Path
                                    ToFolder = 'Inbox'
                                    Status = 'AlreadyMoved'
                                    Strategy = 'Skipped'
                                    Tried = $moveResult.TriedUris
                                })
                            }
                            else {
                                Write-Warning "  Failed to move message: $($moveResult.Error)"
                                
                                $failureRecord = [PSCustomObject]@{ 
                                    MessageId = $m.Id
                                    FromFolder = $c.Path
                                    ToFolder = 'Inbox'
                                    Status = 'Failed'
                                    Error = $moveResult.Error
                                    Tried = $moveResult.TriedUris
                                }
                            
                                if ($Diagnostics) {
                                    $diag = Get-MessageDiagnostics -mid $m.Id -folderId $c.Id
                                    $failureRecord | Add-Member -NotePropertyName 'Diagnostics' `
                                        -NotePropertyValue ($diag | ConvertTo-Json -Depth 3 -Compress)
                                }
                                
                                [void]$auditFailed.Add($failureRecord)
                            }
                        }
                    }
                    catch {
                        Write-Warning "  Exception moving message: $($_.Exception.Message)"
                        [void]$auditFailed.Add([PSCustomObject]@{ 
                            MessageId = $m.Id
                            FromFolder = $c.Path
                            ToFolder = 'Inbox'
                            Status = 'Failed'
                            Error = $_.Exception.Message
                        })
                    }
                }
            }

            # Delete the folder using fresh ID
            if ($DryRun) { 
                Write-Host "  (DryRun) Would delete folder: $($c.Path)" -ForegroundColor Gray
                continue 
            }
            
            try {
                $encFolderDel = Get-EncodedGraphId -Value $freshId
                $delUri = "users/$($TargetUser.Id)/mailFolders/$encFolderDel"
                Invoke-GraphWithRetry -Method DELETE -Uri $delUri -ErrorAction Stop
                Write-LogMessage -Message "Deleted folder: $($c.Path)" -Level 'SUCCESS'
                [void]$auditDeleted.Add([PSCustomObject]@{ 
                    FolderPath = $c.Path
                    FolderId = $freshId
                    Status = 'Deleted'
                    Timestamp = (Get-Date).ToString('o')
                })
            }
            catch {
                # Check if folder was already deleted (NotFound)
                if ($_.Exception.Message -match 'NotFound') {
                    Write-LogMessage -Message "Folder already deleted: $($c.Path)" -Level 'INFO'
                    [void]$auditDeleted.Add([PSCustomObject]@{ 
                        FolderPath = $c.Path
                        FolderId = $freshId
                        Status = 'AlreadyDeleted'
                        Timestamp = (Get-Date).ToString('o')
                    })
                }
                else {
                    Write-LogMessage -Message "Failed to delete folder: $($_.Exception.Message)" -Level 'ERROR'
                    Add-FailureAudit -type 'DeleteFolder' -id $freshId -msg $_.Exception.Message -tried $delUri
                }
            }
        }

        Write-ProgressUpdate -Activity "Processing folder removals" -Completed -Id 2
        Write-LogMessage -Message "Removal pass complete" -Level 'SUCCESS'
    }
}

# Export folder structure
Write-LogMessage -Message "Exporting folder structure..." -Level 'INFO'
$defaultOutDir = Join-Path $repoRoot 'data/output/directory_structures'
$outDir = if ($OutputDir) { $OutputDir } else { $defaultOutDir }

if (!(Test-Path -Path $outDir)) { 
    New-Item -ItemType Directory -Path $outDir -Force | Out-Null 
}

$mailboxSafe = Get-SafeFilename -Name $MailboxEmail
$ts = Get-Date -Format 'yyyyMMdd_HHmm'
$baseName = "$($mailboxSafe)_directory_structure_$ts"

try {
    if ($Format -eq 'json') {
        $outPath = Join-Path $outDir "$baseName.json"
        if ((Test-Path $outPath) -and -not $Overwrite) {
            Write-LogMessage -Message "Output file exists and -Overwrite not specified: $outPath" -Level 'WARN'
            $outPath = Join-Path $outDir "$($baseName)_$(Get-Random -Minimum 1000 -Maximum 9999).json"
        }
        $folders | ConvertTo-Json -Depth 4 | Set-Content -Path $outPath -Encoding UTF8 -Force
    } 
    else {
        $outPath = Join-Path $outDir "$baseName.csv"
        if ((Test-Path $outPath) -and -not $Overwrite) {
            Write-LogMessage -Message "Output file exists and -Overwrite not specified: $outPath" -Level 'WARN'
            $outPath = Join-Path $outDir "$($baseName)_$(Get-Random -Minimum 1000 -Maximum 9999).csv"
        }
        $folders | Export-Csv -Path $outPath -NoTypeInformation -Encoding UTF8 -Force
    }
    Write-LogMessage -Message "Exported directory structure to: $outPath" -Level 'SUCCESS'
    Write-Host "`n✓ Exported to: $outPath" -ForegroundColor Green

    # Save audit logs if ensure operations were performed
    if ($EnsureFromFile) {
        $logDir = Join-Path -Path $outDir -ChildPath 'ensure_logs'
        if (!(Test-Path -Path $logDir)) { 
            New-Item -ItemType Directory -Path $logDir -Force | Out-Null 
        }
        
        $auditLogPath = Join-Path -Path $logDir -ChildPath "$($mailboxSafe)_ensure_$ts.json"
        $audit = [PSCustomObject]@{
            Mailbox = $MailboxEmail
            Timestamp = $ts
            EnsureFile = (Resolve-Path -Path $EnsureFromFile -ErrorAction SilentlyContinue)
            DryRun = $DryRun.IsPresent
            Force = $Force.IsPresent
            Summary = [PSCustomObject]@{
                FoldersCreated = $auditCreated.Count
                FoldersDeleted = $auditDeleted.Count
                MessagesMove = $auditMoved.Count
                OperationsFailed = $auditFailed.Count
            }
        }
        $audit | ConvertTo-Json -Depth 4 | Set-Content -Path $auditLogPath -Encoding UTF8 -Force
        Write-LogMessage -Message "Wrote ensure audit log to: $auditLogPath" -Level 'SUCCESS'

        # Export audit CSVs
        try {
            if ($auditCreated.Count -gt 0) { 
                $csvPath = Join-Path $logDir "$($mailboxSafe)_created_$ts.csv"
                $auditCreated | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8 -Force
                Write-LogMessage -Message "Created folders audit: $csvPath" -Level 'INFO'
            }
            if ($auditMoved.Count -gt 0) { 
                $csvPath = Join-Path $logDir "$($mailboxSafe)_moved_$ts.csv"
                $auditMoved | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8 -Force
                Write-LogMessage -Message "Moved messages audit: $csvPath" -Level 'INFO'
            }
            if ($auditDeleted.Count -gt 0) { 
                $csvPath = Join-Path $logDir "$($mailboxSafe)_deleted_$ts.csv"
                $auditDeleted | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8 -Force
                Write-LogMessage -Message "Deleted folders audit: $csvPath" -Level 'INFO'
            }
            if ($auditFailed.Count -gt 0) { 
                $csvPath = Join-Path $logDir "$($mailboxSafe)_failed_$ts.csv"
                $auditFailed | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8 -Force
                Write-LogMessage -Message "Failed operations audit: $csvPath" -Level 'WARN'
            }
            Write-Host "✓ Audit logs saved to: $logDir" -ForegroundColor Green
        } 
        catch { 
            Write-LogMessage -Message "Failed to write audit CSVs: $($_.Exception.Message)" -Level 'ERROR'
        }
        
        # Display summary
        $actualDeleted = ($auditDeleted | Where-Object { $_.Status -eq 'Deleted' }).Count
        $alreadyDeleted = ($auditDeleted | Where-Object { $_.Status -eq 'AlreadyDeleted' }).Count
        $actualMoved = ($auditMoved | Where-Object { $_.Status -eq 'Moved' }).Count
        $alreadyMoved = ($auditMoved | Where-Object { $_.Status -eq 'AlreadyMoved' }).Count
        
        Write-Host "`n=== Ensure Operations Summary ===" -ForegroundColor Cyan
        Write-Host "Folders created:     $($auditCreated.Count)" -ForegroundColor $(if ($auditCreated.Count -gt 0) { 'Green' } else { 'Gray' })
        Write-Host "Folders deleted:     $actualDeleted" -ForegroundColor $(if ($actualDeleted -gt 0) { 'Yellow' } else { 'Gray' })
        if ($alreadyDeleted -gt 0) {
            Write-Host "  (already deleted): $alreadyDeleted" -ForegroundColor Gray
        }
        Write-Host "Messages moved:      $actualMoved" -ForegroundColor $(if ($actualMoved -gt 0) { 'Cyan' } else { 'Gray' })
        if ($alreadyMoved -gt 0) {
            Write-Host "  (already moved):   $alreadyMoved" -ForegroundColor Gray
        }
        Write-Host "Failed operations:   $($auditFailed.Count)" -ForegroundColor $(if ($auditFailed.Count -gt 0) { 'Red' } else { 'Gray' })
        
        if ($alreadyDeleted -gt 0 -or $alreadyMoved -gt 0) {
            Write-Host "`nNote: Some items were already deleted/moved (possibly by parent folder deletion)" -ForegroundColor Gray
        }
    }
} 
catch {
    Write-LogMessage -Message "Failed to write output: $($_.Exception.Message)" -Level 'ERROR'
    Write-Error "Failed to write output: $($_.Exception.Message)"
}

Write-LogMessage -Message "=== Script completed ===" -Level 'INFO'
Write-Host "`nScript completed. Log file: $logPath" -ForegroundColor Gray
