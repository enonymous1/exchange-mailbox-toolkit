<#
.SYNOPSIS
    Interactively sorts messages from the Inbox into specific target folders.

.DESCRIPTION
    1. Connects to the mailbox via Microsoft Graph.
    2. Reads a list of valid target folders from a text file and resolves their IDs.
    3. Retrieves messages from the Inbox.
    4. For each message, presents a summary to the user and asks for a destination folder.
    5. Supports partial name matching for quick sorting.

.PARAMETER MailboxEmail
    The email address of the target mailbox.

.PARAMETER FolderListFile
    Path to a text file containing the valid list of destination folders.
#>
#requires -Version 7.2

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true, HelpMessage = "The email address of the target mailbox.")]
    [string]$MailboxEmail,

    [Parameter(Mandatory = $true, HelpMessage = "Path to a text file containing list of valid destination folders.")]
    [string]$FolderListFile
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

function Convert-HtmlToText {
	param ([string]$HtmlContent)
	if ([string]::IsNullOrWhiteSpace($HtmlContent)) { return "" }
	$text = $HtmlContent -replace '<br[^>]*>', "`r`n"
	$text = $text -replace '<[^>]+>', ""
	$text = $text -replace '&nbsp;', ' ' 
	$text = $text -replace '&amp;', '&'
	$text = $text -replace '&lt;', '<'
	return $text -replace "(?m)^\s*\r?\n", ""
}

# --- Load Utilities ---
try {
    . "$PSScriptRoot/../Utils/Initialize-Environment.ps1"
    . "$PSScriptRoot/../Utils/LoggingHelpers.ps1"
} catch {
    Write-Error "Failed to load Utils. Ensure '..\Utils\Initialize-Environment.ps1' exists."
    return
}

# --- Initialize ---
$EnvStatus = Connect-ExchangeEnvironment -GraphScopes "User.Read.All","Mail.ReadWrite.Shared"
if (-not $EnvStatus.Graph.Connected) { return }

# --- Logging ---
$repoRoot = Resolve-Path -Path (Join-Path $PSScriptRoot '../..') | Select-Object -ExpandProperty Path
$logDir = Join-Path $repoRoot 'logs/Sort-Inbox'
if (!(Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
$timestamp = Get-Date -Format "yyyyMMdd_HHmm"
$mailboxSafe = Get-SafeFilename -Subject $MailboxEmail
Set-LogFile -Path (Join-Path $logDir "$($mailboxSafe)_sort_inbox_$timestamp.log")
Write-LogMessage -Level INFO -Message "Starting Sort-Inbox for $MailboxEmail"

# --- Get User ---
$TargetUser = Get-MgUser -Filter "mail eq '$MailboxEmail'" -ConsistencyLevel eventual -CountVariable c -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $TargetUser) {
    $TargetUser = Get-MgUser -Filter "proxyAddresses/any(p:p eq 'smtp:$MailboxEmail')" -ConsistencyLevel eventual -CountVariable c -ErrorAction SilentlyContinue | Select-Object -First 1
}
if (-not $TargetUser) {
    Write-Error "User '$MailboxEmail' not found."
    return
}
Write-LogMessage -Level INFO -Message "Target User: $($TargetUser.DisplayName)"

# --- Resolve Target Folders ---
if (-not (Test-Path $FolderListFile)) { Write-Error "Folder list not found"; return }
$validPaths = Get-Content $FolderListFile | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.Trim() }
Write-Host "Resolving $($validPaths.Count) target folders..." -ForegroundColor Cyan

$folderMap = @{} # Name -> ID mapping
$flatList = @()  # List for efficient searching

# Helper to find folder ID by path (using cache or Graph)
# For simplicity in this interactive script, we'll traverse Graph once to map IDs, 
# or look them up on demand? Traversing once is safer for "valid targets".
# However, mapping arbitrary paths like 'Inbox\A\B' requires recursive lookups.

function Resolve-FolderPathToId {
    param ($Path)
    # Basic caching logic could live here, but for now we do direct resolution
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
        if ($found) { $currentId = $found.Id } else { return $null }
    }
    return $currentId
}

Write-Host "Pre-verifying folders from list (this might take a minute)..." -ForegroundColor Cyan
foreach ($path in $validPaths) {
    $id = Resolve-FolderPathToId -Path $path
    if ($id) {
        $folderMap[$path] = $id
        $flatList += $path
        Write-Verbose "Mapped: $path"
    } else {
        Write-Warning "Target folder not found in mailbox: $path"
    }
}
Write-Host "Mapped $($folderMap.Count) valid destination folders." -ForegroundColor Green

# --- Fetch Inbox ---
$Inbox = Get-MgUserMailFolder -UserId $TargetUser.Id -Filter "displayName eq 'Inbox'" -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $Inbox) { Write-Error "Inbox not found"; return }

Write-Host "Fetching Inbox messages..." -ForegroundColor Cyan
# Fetch Top 50 to start
$messages = Get-MgUserMailFolderMessage -UserId $TargetUser.Id -MailFolderId $Inbox.Id -Top 50 -All:$false -ErrorAction SilentlyContinue

if ($messages.Count -eq 0) {
    Write-Host "Inbox is empty!" -ForegroundColor Green
    return
}

# --- Interactive Loop ---
$processedCount = 0
foreach ($msg in $messages) {
    Clear-Host
    Write-Host "--- MESSAGE [$($processedCount + 1)/$($messages.Count)] ---" -ForegroundColor Yellow
    Write-Host "From:    " -NoNewline; Write-Host $msg.From.EmailAddress.Address -ForegroundColor White
    Write-Host "Subject: " -NoNewline; Write-Host $msg.Subject -ForegroundColor Cyan
    Write-Host "Received:" -NoNewline; Write-Host $msg.ReceivedDateTime -ForegroundColor Gray
    
    $preview = "No text content"
    if ($msg.BodyPreview) { $preview = $msg.BodyPreview }
    elseif ($msg.Body.Content) { $preview = (Convert-HtmlToText $msg.Body.Content).Substring(0, [math]::Min($msg.Body.Content.Length, 200)) }
    
    Write-Host "Preview: " -NoNewline; Write-Host "$preview..." -ForegroundColor DarkGray
    Write-Host "----------------------------------------------------" -ForegroundColor Yellow

    # Loop for valid selection
    while ($true) {
        $inputSearch = Read-Host "[Type folder name] or [S]kip or [Q]uit"
        
        if ($inputSearch -eq 'Q') {
            Write-Host "Quitting..."
            return
        }
        if ($inputSearch -eq 'S' -or [string]::IsNullOrWhiteSpace($inputSearch)) {
            Write-Host "Skipping..." -ForegroundColor Gray
            break
        }

        # Search Logic
        $foundFolders = $folderMap.Keys | Where-Object { $_ -like "*$inputSearch*" }
        
        if ($foundFolders.Count -eq 0) {
            Write-Warning "No folders match '$inputSearch'."
        }
        elseif ($foundFolders.Count -eq 1) {
            $targetName = $foundFolders[0]
            $confirm = Read-Host "Move to '$targetName'? (Y/N)"
            if ($confirm -eq 'Y' -or $confirm -eq '') {
                Try {
                    Move-MgUserMailFolderMessage -UserId $TargetUser.Id -MailFolderId $Inbox.Id -MessageId $msg.Id -DestinationId $folderMap[$targetName] -ErrorAction Stop
                    Write-Host "Moved!" -ForegroundColor Green
                    Write-LogMessage -Level INFO -Message "Moved '$($msg.Subject)' to '$targetName'"
                    break
                } Catch {
                    Write-Error "Failed to move: $_"
                }
            }
        }
        else {
            Write-Host "Multiple matches found:" -ForegroundColor Magenta
            $i = 0
            foreach ($m in $foundFolders) {
                Write-Host "  [$i] $m"
                $i++
            }
            $select = Read-Host "Select number (or Enter to cancel)"
            
            if ($select -match '^\d+$' -and $select -ge 0 -and $select -lt $foundFolders.Count) {
                $targetName = $foundFolders[$select]
                Try {
                    Move-MgUserMailFolderMessage -UserId $TargetUser.Id -MailFolderId $Inbox.Id -MessageId $msg.Id -DestinationId $folderMap[$targetName] -ErrorAction Stop
                    Write-Host "Moved!" -ForegroundColor Green
                    Write-LogMessage -Level INFO -Message "Moved '$($msg.Subject)' to '$targetName'"
                    break
                } Catch {
                    Write-Error "Failed to move: $_"
                }
            }
        }
    }
    $processedCount++
}

Write-Host "Batch complete." -ForegroundColor Green
