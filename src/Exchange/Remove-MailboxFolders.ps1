<#
.SYNOPSIS
    Removes specified folders from a mailbox after flattening (moving) all contents to the Inbox.

.DESCRIPTION
    1. Connects to the mailbox via Microsoft Graph.
    2. Reads a list of folder paths to remove.
    3. For each target folder:
       a. Recursively moves ALL items (from the folder and its subfolders) to the Inbox.
       b. Deletes the empty folder tree.
    4. Protects standard directories (Inbox, Sent Items, etc.) from accidental deletion.

.PARAMETER MailboxEmail
    The email address of the target mailbox.

.PARAMETER FolderListFile
    Path to a text file containing the specific folder paths to remove.

.PARAMETER DestinationFolderName
    The name of the folder where items should be moved to. Defaults to 'Inbox'.
#>
#requires -Version 7.2

[CmdletBinding(SupportsShouldProcess=$true)]
param (
    [Parameter(Mandatory = $true, HelpMessage = "The email address of the target mailbox.")]
    [string]$MailboxEmail,

    [Parameter(Mandatory = $true, HelpMessage = "Path to a text file containing paths of folders to delete.")]
    [string]$FolderListFile,

    [string]$DestinationFolderName = 'Inbox'
)

# --- Helper Functions ---
function Get-SafeFilename {
    param ([string]$Subject)
    if ($null -eq $Subject) { $Subject = '' }
    $safe = $Subject -replace '[\\/:*?"<>|]', '_'
    $safe = $safe -replace "[\r\n\t]", ' '
    if ($safe.Length -gt 50) { $safe = $safe.Substring(0,50) }
    return $safe.Trim()
}

# --- Load Utilities ---
try {
    . "$PSScriptRoot/../Utils/Initialize-Environment.ps1"
    . "$PSScriptRoot/../Utils/LoggingHelpers.ps1"
} catch {
    Write-Error "Failed to load Utils. Ensure '..\Utils\Initialize-Environment.ps1' and 'LoggingHelpers.ps1' exist."
    return
}

# --- Initialize Environment ---
$EnvStatus = Connect-ExchangeEnvironment -GraphScopes "User.Read.All","Mail.Read.Shared","Mail.ReadWrite.Shared"
if (-not $EnvStatus.Graph.Connected) {
    Write-Error "Failed to connect to Microsoft Graph: $($EnvStatus.Graph.Message)"
    return
}

# --- Setup Logging ---
$repoRoot = Resolve-Path -Path (Join-Path $PSScriptRoot '../..') | Select-Object -ExpandProperty Path
$logDir = Join-Path $repoRoot 'logs/Remove-MailboxFolders'
if (!(Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }

$mailboxSafe = Get-SafeFilename -Subject $MailboxEmail
$timestamp = Get-Date -Format "yyyyMMdd_HHmm"
$logName = "$($mailboxSafe)_cleanup_$timestamp.log"

Set-LogFile -Path (Join-Path $logDir $logName)
Write-LogMessage -Level INFO -Message "Starting Remove-MailboxFolders for $MailboxEmail"

# --- Validate Input ---
if (-not (Test-Path $FolderListFile)) {
    Write-Error "Input file not found: $FolderListFile"
    return
}

# --- Get User ---
$TargetUser = Get-MgUser -Filter "mail eq '$MailboxEmail'" -ConsistencyLevel eventual -CountVariable c -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $TargetUser) {
    $TargetUser = Get-MgUser -Filter "proxyAddresses/any(p:p eq 'smtp:$MailboxEmail')" -ConsistencyLevel eventual -CountVariable c -ErrorAction SilentlyContinue | Select-Object -First 1
}

if (-not $TargetUser) {
    Write-Error "User '$MailboxEmail' not found."
    return
}
Write-LogMessage -Level INFO -Message "Target User: $($TargetUser.DisplayName) ($($TargetUser.Id))"

# --- Get Destination Folder ---
$DestFolder = Get-MgUserMailFolder -UserId $TargetUser.Id -Filter "displayName eq '$DestinationFolderName'" -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $DestFolder) {
    Write-Error "Destination folder '$DestinationFolderName' not found."
    return
}
$DestId = $DestFolder.Id
Write-LogMessage -Level INFO -Message "Destination Folder: $($DestFolder.DisplayName) ($DestId)"

# --- Protection List ---
$ProtectedFolders = @("Inbox", "Drafts", "Sent Items", "Deleted Items", "Junk Email", "Archive", "Outbox", "Conversation History")

# --- Logic: Flatten and Delete ---
function Resolve-FolderPathToId {
    param ($Path)
    $parts = $Path -split '\\'
    $currentId = $null
    
    foreach ($part in $parts) {
        $params = @{ UserId = $TargetUser.Id; Filter = "displayName eq '$part'" }
        if ($currentId) {
            $params["MailFolderId"] = $currentId
            $found = Get-MgUserMailFolderChildFolder @params -ErrorAction SilentlyContinue | Select-Object -First 1
        } else {
            $found = Get-MgUserMailFolder @params -ErrorAction SilentlyContinue | Select-Object -First 1
        }

        if ($found) {
            $currentId = $found.Id
        } else {
            return $null
        }
    }
    return $currentId
}

function Move-FolderContentsRecursive {
    param ($FolderId, $FolderPath)

    # 1. Move Items in Current Folder
    $messages = Get-MgUserMailFolderMessage -UserId $TargetUser.Id -MailFolderId $FolderId -PageSize 100 -All -ErrorAction SilentlyContinue
    $count = 0
    foreach ($msg in $messages) {
        try {
            Move-MgUserMailFolderMessage -UserId $TargetUser.Id -MailFolderId $FolderId -MessageId $msg.Id -DestinationId $DestId -ErrorAction Stop | Out-Null
            $count++
        } catch {
            Write-LogMessage -Level ERROR -Message "Failed to move message $($msg.Id) from $FolderPath. Error: $($_.Exception.Message)"
        }
    }
    if ($count -gt 0) {
        Write-LogMessage -Level INFO -Message "Moved $count items from '$FolderPath' to '$DestinationFolderName'."
    }

    # 2. Recurse into Children
    $children = Get-MgUserMailFolderChildFolder -UserId $TargetUser.Id -MailFolderId $FolderId -PageSize 100 -All -ErrorAction SilentlyContinue
    foreach ($child in $children) {
        Move-FolderContentsRecursive -FolderId $child.Id -FolderPath "$FolderPath\$($child.DisplayName)"
    }
}

$pathsToRemove = Get-Content -Path $FolderListFile | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.Trim() }

# Sort paths by length descending to delete children before parents
# This optimizes API calls, though strict recursion handles it too.
$pathsToRemove = $pathsToRemove | Sort-Object { $_.Length } -Descending

foreach ($path in $pathsToRemove) {
    Write-Host "Processing: $path" -ForegroundColor Cyan
    Write-LogMessage -Level INFO -Message "Processing target: $path"

    # Safety Check
    $isProtected = $false
    foreach ($prot in $ProtectedFolders) {
        if ($path -eq $prot) { $isProtected = $true }
    }
    if ($isProtected) {
        Write-Warning "Skipping protected folder: $path"
        Write-LogMessage -Level WARN -Message "Skipped protected folder: $path"
        continue
    }

    # Resolve ID
    $targetId = Resolve-FolderPathToId -Path $path
    if (-not $targetId) {
        Write-Warning "Folder not found: $path (It may have been deleted already)"
        Write-LogMessage -Level WARN -Message "Folder not found: $path"
        continue
    }

    # Move Contents
    Write-Host "  Flattening contents..." -ForegroundColor Gray
    try {
        Move-FolderContentsRecursive -FolderId $targetId -FolderPath $path
    } catch {
        Write-Error "Failed during content flattening for $path. Aborting delete. $_"
        Write-LogMessage -Level ERROR -Message "Flattening failed for $path. Error: $_"
        continue
    }

    # Delete Folder (v2.0.0: gated by ShouldProcess, so -WhatIf / -Confirm work)
    Write-Host "  Deleting folder..." -ForegroundColor Yellow
    if (-not $PSCmdlet.ShouldProcess("$path in $MailboxEmail", 'Delete mailbox folder')) {
        Write-LogMessage -Level INFO -Message "Skipped delete of folder (WhatIf/declined): $path"
        continue
    }
    try {
        Remove-MgUserMailFolder -UserId $TargetUser.Id -MailFolderId $targetId -ErrorAction Stop
        Write-Host "  Success." -ForegroundColor Green
        Write-LogMessage -Level INFO -Message "Successfully deleted folder: $path"
    } catch {
        Write-Error "Failed to delete folder $path. Error: $($_.Exception.Message)"
        Write-LogMessage -Level ERROR -Message "Failed to delete folder $path. Error: $($_.Exception.Message)"
    }
}

Write-Host "Operation Complete." -ForegroundColor Green
