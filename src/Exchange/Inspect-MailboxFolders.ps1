<#
.SYNOPSIS
    Inspects a specific list of folders within a mailbox and exports detailed information.
    Can be used to verify existence and gather stats (ItemCount) for a targeted list of folders.

.DESCRIPTION
    1. Connects to Microsoft Graph.
    2. Retrieves the full folder structure of the target mailbox to ensure accurate path matching.
    3. Reads a provided text file containing a list of folder paths (one per line).
    4. Matches the input paths against the actual mailbox structure.
    5. Exports a report detailing the status (Found/Missing), ID, and ItemCounts for each requested folder.

.PARAMETER MailboxEmail
    The email address of the target mailbox.

.PARAMETER FolderListFile
    Path to a text file containing the list of folder paths to inspect.
#>
#requires -Version 7.2

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true, HelpMessage = "The email address of the mailbox to inspect.")]
    [string]$MailboxEmail,

    [Parameter(Mandatory = $true, HelpMessage = "Path to a text file containing list of folder paths to check.")]
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

# --- Load Utilities ---
try {
    . "$PSScriptRoot/../Utils/Initialize-Environment.ps1"
    . "$PSScriptRoot/../Utils/LoggingHelpers.ps1"
} catch {
    Write-Error "Failed to load Utils. Ensure '..\Utils\Initialize-Environment.ps1' and 'LoggingHelpers.ps1' exist."
    return
}

# --- Initialize Environment ---
$EnvStatus = Connect-ExchangeEnvironment -GraphScopes "User.Read.All","Mail.Read.Shared"
if (-not $EnvStatus.Graph.Connected) {
    Write-Error "Failed to connect to Microsoft Graph: $($EnvStatus.Graph.Message)"
    return
}

# --- Setup Logging ---
$repoRoot = Resolve-Path -Path (Join-Path $PSScriptRoot '../..') | Select-Object -ExpandProperty Path
$logDir = Join-Path $repoRoot 'logs/Inspect-MailboxFolders'
if (!(Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }

$mailboxSafe = Get-SafeFilename -Subject $MailboxEmail
$timestamp = Get-Date -Format "yyyyMMdd_HHmm"
$logName = "$($mailboxSafe)_inspect_folders_$timestamp.log"

Set-LogFile -Path (Join-Path $logDir $logName)
Write-LogMessage -Level INFO -Message "Starting Inspect-MailboxFolders for $MailboxEmail"

# --- Validate Input File ---
if (-not (Test-Path $FolderListFile)) {
    Write-LogMessage -Level ERROR -Message "Input file not found: $FolderListFile"
    Write-Error "Input file not found: $FolderListFile"
    return
}

$targetPaths = Get-Content -Path $FolderListFile | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.Trim() }
Write-LogMessage -Level INFO -Message "Loaded $($targetPaths.Count) paths to inspect from file."

# --- Get User ---
Write-LogMessage -Level INFO -Message "Looking up user: $MailboxEmail"
$TargetUser = Get-MgUser -Filter "mail eq '$MailboxEmail'" -ConsistencyLevel eventual -CountVariable c -ErrorAction SilentlyContinue | Select-Object -First 1

if (-not $TargetUser) {
    Write-LogMessage -Level WARN -Message "Primary lookup failed. Checking proxy addresses..."
    $TargetUser = Get-MgUser -Filter "proxyAddresses/any(p:p eq 'smtp:$MailboxEmail')" -ConsistencyLevel eventual -CountVariable c -ErrorAction SilentlyContinue | Select-Object -First 1
}

if (-not $TargetUser) {
    Write-LogMessage -Level ERROR -Message "CRITICAL: User '$MailboxEmail' not found."
    Write-Error "CRITICAL: User '$MailboxEmail' not found."
    return
}
Write-LogMessage -Level INFO -Message "Found User: $($TargetUser.DisplayName) ($($TargetUser.Id))"

# --- Build Actual Folder Map ---
# We fetch all folders to build a reliable Path -> ID map. This avoids ambiguity with duplicates at different levels.
Write-Host "Fetching actual mailbox structure..." -ForegroundColor Cyan
Write-LogMessage -Level INFO -Message "Fetching actual mailbox structure..."

$actualFoldersMap = @{} # Key: Path, Value: FolderObject

function Get-MailboxFoldersRecursive {
    param (
        [string]$ParentFolderId,
        [string]$ParentFolderPath
    )

    $params = @{
        UserId = $TargetUser.Id
        PageSize = 999
        All = $true
    }
    
    if ($ParentFolderId) {
        $params.MailFolderId = $ParentFolderId
        $folders = Get-MgUserMailFolderChildFolder @params -ErrorAction SilentlyContinue
    } else {
        $folders = Get-MgUserMailFolder @params -ErrorAction SilentlyContinue
    }

    foreach ($folder in $folders) {
        if ($ParentFolderPath) {
            $currentPath = "$ParentFolderPath\$($folder.DisplayName)"
        } else {
            $currentPath = $folder.DisplayName
        }

        # Store in map (Case-insensitive key by default in PowerShell hashtables)
        if (-not $actualFoldersMap.ContainsKey($currentPath)) {
            $actualFoldersMap[$currentPath] = $folder
        }

        if ($folder.ChildFolderCount -gt 0) {
            Get-MailboxFoldersRecursive -ParentFolderId $folder.Id -ParentFolderPath $currentPath
        }
    }
}

Get-MailboxFoldersRecursive -ParentFolderId $null -ParentFolderPath $null
Write-LogMessage -Level INFO -Message "Mapped $($actualFoldersMap.Count) actual folders in mailbox."

# --- Inspect Requested Folders ---
Write-Host "Inspecting dependencies..." -ForegroundColor Cyan
$report = [System.Collections.Generic.List[PSObject]]::new()

foreach ($path in $targetPaths) {
    $exists = $actualFoldersMap.ContainsKey($path)
    $folderObj = $null
    
    if ($exists) {
        $folderObj = $actualFoldersMap[$path]
        $status = "Found"
        $id = $folderObj.Id
        $itemCount = $folderObj.TotalItemCount
        $childCount = $folderObj.ChildFolderCount
    } else {
        $status = "Missing"
        $id = $null
        $itemCount = 0
        $childCount = 0
    }

    $row = [PSCustomObject]@{
        RequestedPath = $path
        Status        = $status
        FolderId      = $id
        TotalItems    = $itemCount
        ChildFolders  = $childCount
    }
    $report.Add($row)
    
    if ($exists) {
        Write-Verbose "Found: $path (Items: $itemCount)"
    } else {
        Write-Verbose "Missing: $path"
    }
}

# --- Export Report ---
# Output to directory
$dataOutDir = Join-Path $repoRoot "data/output/$mailboxSafe/FolderStructure"
if (!(Test-Path $dataOutDir)) { New-Item -ItemType Directory -Path $dataOutDir -Force | Out-Null }

$csvPath = Join-Path $dataOutDir "$($mailboxSafe)_InspectionReport_$timestamp.csv"

try {
    $report | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
    Write-LogMessage -Level INFO -Message "Inspection report exported to: $csvPath"
    Write-Host "Inspection completed. Report saved to: $csvPath" -ForegroundColor Green
    
    # Summary
    $foundCount = ($report | Where-Object { $_.Status -eq 'Found' }).Count
    $missingCount = ($report | Where-Object { $_.Status -eq 'Missing' }).Count
    Write-Host "Summary: Found $foundCount, Missing $missingCount" -ForegroundColor Yellow
} catch {
    Write-LogMessage -Level ERROR -Message "Failed to export CSV: $($_.Exception.Message)"
    Write-Error "Failed to export CSV: $($_.Exception.Message)"
}
