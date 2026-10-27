<#
.SYNOPSIS
Bulk move messages based on an input file map (MessageID -> TargetFolder).

.DESCRIPTION
Reads a CSV or text file containing a mapping of MessageIDs to Target Folder Paths.
For each entry, it ensures the target folder exists (creating it if necessary) and moves the message.

.PARAMETER MailboxEmail
The target mailbox email address.

.PARAMETER InputFile
Path to the input file. Expects a CSV format.
Supported formats:
1. Headers: "MessageId", "TargetFolder" (case-insensitive)
2. No Headers: Column 1 is MessageId, Column 2 is TargetFolder Path.

.PARAMETER DryRun
If set, only simulates the moves.

.EXAMPLE
.\Import-MoveMessages.ps1 -MailboxEmail "user@domain.com" -InputFile "data/input/move_map.csv"
#>
#requires -Version 7.2
[CmdletBinding(SupportsShouldProcess=$true)]
param (
    [Parameter(Mandatory = $true)]
    [string]$MailboxEmail,

    [Parameter(Mandatory = $true)]
    [string]$InputFile,

    [switch]$DryRun
)

# --- Load Utilities ---
try {
    . "$PSScriptRoot/../Utils/Initialize-Environment.ps1"
    . "$PSScriptRoot/../Utils/LoggingHelpers.ps1"
} catch {
    Write-Error "Failed to load Utils."
    return
}

# --- Initialize ---
# Connect to Graph
$EnvStatus = Connect-ExchangeEnvironment -GraphScopes "User.Read.All","Mail.ReadWrite.Shared","Mail.ReadWrite"
if (-not $EnvStatus.Graph.Connected) {
    Write-Error "Failed to connect to Microsoft Graph."
    return
}

# Setup Logging
$repoRoot = Resolve-Path -Path (Join-Path $PSScriptRoot '../..') | Select-Object -ExpandProperty Path
$logDir = Join-Path $repoRoot 'logs/Import-MoveMessages'
if (!(Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
$logName = "$(Get-Date -Format 'yyyyMMdd_HHmm')_move_messages.log"
Set-LogFile -Path (Join-Path $logDir $logName)

Write-LogMessage -Level INFO -Message "Starting Bulk Move for $MailboxEmail"

# Resolve User
$TargetUser = Get-MgUser -Filter "mail eq '$MailboxEmail'" -ConsistencyLevel eventual -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $TargetUser) {
    $TargetUser = Get-MgUser -Filter "proxyAddresses/any(p:p eq 'smtp:$MailboxEmail')" -ConsistencyLevel eventual -ErrorAction SilentlyContinue | Select-Object -First 1
}
if (!$TargetUser) {
    Write-Error "Mailbox not found: $MailboxEmail"
    Write-LogMessage -Level ERROR -Message "Mailbox not found."
    return
}
Write-LogMessage -Level INFO -Message "Found Mailbox: $($TargetUser.DisplayName) ($($TargetUser.Id))"

# Read Input File
if (!(Test-Path $InputFile)) {
    Write-Error "Input file not found: $InputFile"
    return
}

$rawContent = Get-Content -Path $InputFile
$delimiter = ','
if ($rawContent[0] -match '\t') { $delimiter = "`t" }

try {
    # Check for headers
    $firstLine = $rawContent[0]
    $hasHeaders = ($firstLine -match 'MessageId' -and $firstLine -match 'TargetFolder')

    if ($hasHeaders) {
        $SourceData = Import-Csv -Path $InputFile -Delimiter $delimiter
    } else {
        $SourceData = Import-Csv -Path $InputFile -Header 'MessageId','TargetFolder' -Delimiter $delimiter
    }
} catch {
    Write-Error "Failed to parse CSV: $($_.Exception.Message)"
    return
}

Write-LogMessage -Level INFO -Message "Loaded $($SourceData.Count) items to process."

# Cache for Folder Ids to avoid repeated lookups
$FolderCache = @{} # Key: Path, Value: ID

function Get-FolderIdByPath {
    param ([string]$Path)

    # Normalize Path
    $Path = $Path -replace '/', '\'
    $Path = $Path.Trim('\')
    
    if ($FolderCache.ContainsKey($Path)) { return $FolderCache[$Path] }

    $segments = $Path -split '\\'
    $currentParentId = $null # Null implies MsgFolderRoot

    foreach ($seg in $segments) {
        # Find folder in current parent
        if ($null -eq $currentParentId) {
             # Look at root (Top level)
             # Note: Using ChildFolders of Root is safer than "Get-MgUserMailFolder" which gets top level
             $found = Get-MgUserMailFolder -UserId $TargetUser.Id -Filter "displayName eq '$seg'" -Top 1 -ErrorAction SilentlyContinue
        } else {
             $found = Get-MgUserMailFolderChildFolder -UserId $TargetUser.Id -MailFolderId $currentParentId -Filter "displayName eq '$seg'" -Top 1 -ErrorAction SilentlyContinue
        }

        if ($found) {
            $currentParentId = $found.Id
        } else {
            # Fail if folder missing (User requested no auto-creation)
            Write-LogMessage -Level WARN -Message "Folder segment '$seg' missing in path '$Path'. Auto-creation is disabled."
            # Throwing here will be caught by the calling try/catch block
            throw "Folder path not found: $Path"
        }
    }

    if (!$DryRun) { $FolderCache[$Path] = $currentParentId }
    return $currentParentId
}

# Process Messages
$successCount = 0
$failCount = 0

foreach ($row in $SourceData) {
    $mid = $row.MessageId
    $folderPath = $row.TargetFolder

    if ([string]::IsNullOrWhiteSpace($mid) -or [string]::IsNullOrWhiteSpace($folderPath)) {
        continue
    }

    try {
        if ($DryRun) {
            Write-Host "[DryRun] Move Message $mid -> $folderPath"
            $successCount++
            continue
        }

        # 1. Ensure/Get Target Folder ID
        try {
            $destId = Get-FolderIdByPath -Path $folderPath
        } catch {
            Write-LogMessage -Level ERROR -Message "Skipping message ${mid}: result $($_.Exception.Message)"
            $failCount++
            continue
        }

        # 2. Move Message (v2.0.0: gated by ShouldProcess, so -WhatIf works)
        if (-not $PSCmdlet.ShouldProcess($mid, "Move message to '$folderPath'")) { continue }
        Write-LogMessage -Level INFO -Message "Moving $mid to '$folderPath'..."
        Move-MgUserMessage -UserId $TargetUser.Id -MessageId $mid -DestinationId $destId -ErrorAction Stop
        $successCount++

    } catch {
        Write-LogMessage -Level ERROR -Message "Failed to move message ${mid}: $($_.Exception.Message)"
        $failCount++
    }
}

Write-LogMessage -Level INFO -Message "Complete. Success: $successCount, Failed: $failCount"
Write-Host "Done. Check logs at $(Join-Path $logDir $logName)"
