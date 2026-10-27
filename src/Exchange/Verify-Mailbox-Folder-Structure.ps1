<#
.SYNOPSIS
    Retrieves and exports the folder structure of a specified mailbox.
.DESCRIPTION
    Connects to Microsoft Graph, recursively fetches the folder hierarchy for the target mailbox,
    and optionally exports the structure to a CSV file in the data output directory.
#>
#requires -Version 7.2

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true, HelpMessage = "The email address of the mailbox to inspect (e.g. user@contoso.com).")]
    [string]$MailboxEmail,

    [Parameter(Mandatory = $false, HelpMessage = "Path to a text file containing a list of expected folder paths (one per line) to verify against.")]
    [string]$ExpectedStructureFile
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
# Connect to Graph (User.Read.All and Mail.Read.Shared required)
$EnvStatus = Connect-ExchangeEnvironment -GraphScopes "User.Read.All","Mail.Read.Shared"
if (-not $EnvStatus.Graph.Connected) {
    Write-Error "Failed to connect to Microsoft Graph: $($EnvStatus.Graph.Message)"
    return
}

# --- Setup Logging ---
$repoRoot = Resolve-Path -Path (Join-Path $PSScriptRoot '../..') | Select-Object -ExpandProperty Path
$logDir = Join-Path $repoRoot 'logs/Verify-MailboxFolderStructure'
if (!(Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }

$mailboxSafe = Get-SafeFilename -Subject $MailboxEmail
$timestamp = Get-Date -Format "yyyyMMdd_HHmm"
$logName = "$($mailboxSafe)_folder_structure_$timestamp.log"

Set-LogFile -Path (Join-Path $logDir $logName)
Write-LogMessage -Level INFO -Message "Starting Verify-Mailbox-Folder-Structure for $MailboxEmail"

# --- Main Logic ---
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

$folderList = [System.Collections.Generic.List[PSObject]]::new()

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
        # Top level folders
        $folders = Get-MgUserMailFolder @params -ErrorAction SilentlyContinue
    }

    foreach ($folder in $folders) {
        if ($ParentFolderPath) {
            $currentPath = "$ParentFolderPath\$($folder.DisplayName)"
        } else {
            $currentPath = $folder.DisplayName
        }

        $obj = [PSCustomObject]@{
            DisplayName = $folder.DisplayName
            Path        = $currentPath
            Id          = $folder.Id
            ParentId    = $folder.ParentFolderId
            ChildCount  = $folder.ChildFolderCount
            ItemCount   = $folder.TotalItemCount
        }
        $folderList.Add($obj)
        Write-Verbose "Found: $currentPath"

        if ($folder.ChildFolderCount -gt 0) {
            Get-MailboxFoldersRecursive -ParentFolderId $folder.Id -ParentFolderPath $currentPath
        }
    }
}

Write-LogMessage -Level INFO -Message "Retrieving folder structure..."
Write-Host "Retrieving folder structure. This may take a moment..." -ForegroundColor Cyan

Get-MailboxFoldersRecursive -ParentFolderId $null -ParentFolderPath $null

Write-LogMessage -Level INFO -Message "Total folders found: $($folderList.Count)"
Write-Host "Total folders found: $($folderList.Count)" -ForegroundColor Green

# Export
if ($folderList.Count -gt 0) {
    # Output to directory
    $dataOutDir = Join-Path $repoRoot "data/output/$mailboxSafe/FolderStructure"
    if (!(Test-Path $dataOutDir)) { New-Item -ItemType Directory -Path $dataOutDir -Force | Out-Null }
    
    $csvPath = Join-Path $dataOutDir "$($mailboxSafe)_FolderStructure_$timestamp.csv"
    
    try {
        $folderList | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
        Write-LogMessage -Level INFO -Message "Exported structure to: $csvPath"
        Write-Host "Export completed: $csvPath" -ForegroundColor Green
    } catch {
        Write-LogMessage -Level ERROR -Message "Failed to export CSV: $($_.Exception.Message)"
        Write-Error "Failed to export CSV: $($_.Exception.Message)"
    }
} else {
    Write-LogMessage -Level WARN -Message "No folders returned. This is unexpected for a valid mailbox."
    Write-Warning "No folders found."
}

# --- Comparison Logic ---
if ($ExpectedStructureFile) {
    if (Test-Path $ExpectedStructureFile) {
        Write-LogMessage -Level INFO -Message "Verifying against expected structure file: $ExpectedStructureFile"
        Write-Host "Verifying against expected structure file..." -ForegroundColor Cyan
        
        $expectedFolders = Get-Content -Path $ExpectedStructureFile | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.Trim() }
        
        # --- Check for Missing Folders (Expected but not in Actual) ---
        $actualPaths = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($item in $folderList) {
            $null = $actualPaths.Add($item.Path)
        }

        $missingFolders = [System.Collections.Generic.List[string]]::new()
        $foundFolders = [System.Collections.Generic.List[string]]::new()

        foreach ($expected in $expectedFolders) {
            if ($actualPaths.Contains($expected)) {
                $foundFolders.Add($expected)
            } else {
                $missingFolders.Add($expected)
            }
        }

        # --- Check for Extra Folders (Actual but not in Expected) ---
        $expectedPathsSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($ex in $expectedFolders) { $null = $expectedPathsSet.Add($ex) }
        
        # Standard default Exchange folders to exclude from "Extra" reporting
        $standardFolders = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        # Note: These are top-level standard folders
        $defaults = @("Archive", "Conversation History", "Deleted Items", "Drafts", "Inbox", "Junk Email", "Outbox", "Sent Items")
        foreach ($d in $defaults) { $null = $standardFolders.Add($d) }

        $extraFolders = [System.Collections.Generic.List[string]]::new()
        foreach ($item in $folderList) {
            # Check if expected
            if ($expectedPathsSet.Contains($item.Path)) { continue }
            
            # Check if it's a standard/default folder
            if ($standardFolders.Contains($item.Path)) { continue }

            $extraFolders.Add($item.Path)
        }

        # --- Reporting ---
        Write-Host "Comparison Results:" -ForegroundColor Yellow
        Write-Host "  Expected Total: $($expectedFolders.Count)"
        Write-Host "  Actual Total:   $($folderList.Count)"
        Write-Host "  Matched:        $($foundFolders.Count)"
        Write-Host "  Missing:        $($missingFolders.Count)"
        Write-Host "  Extra:          $($extraFolders.Count)"
        
        Write-LogMessage -Level INFO -Message "Comparison - Expected: $($expectedFolders.Count), Actual: $($folderList.Count), Matched: $($foundFolders.Count), Missing: $($missingFolders.Count), Extra: $($extraFolders.Count)"

        # Report Missing
        if ($missingFolders.Count -gt 0) {
            Write-Host "`nMissing Folders (Expected but not found):" -ForegroundColor Red
            $missingFolders | ForEach-Object { Write-Host "  [-] $_" -ForegroundColor Red }
            
            if ($dataOutDir) {
                $missingLogPath = Join-Path $dataOutDir "$($mailboxSafe)_MissingFolders_$timestamp.txt"
                $missingFolders | Out-File -FilePath $missingLogPath -Encoding UTF8
                Write-LogMessage -Level WARN -Message "Missing folders exported to: $missingLogPath"
                Write-Host "Saved missing list: $missingLogPath" -ForegroundColor Red
            }
        } else {
            Write-Host "`nAll expected folders were found!" -ForegroundColor Green
            Write-LogMessage -Level INFO -Message "All expected folders were found."
        }

        # Report Extra
        if ($extraFolders.Count -gt 0) {
            Write-Host "`nExtra Folders (Found but not in expected list):" -ForegroundColor Magenta
            # Limit console output if too many
            if ($extraFolders.Count -gt 20) {
                $extraFolders | Select-Object -First 20 | ForEach-Object { Write-Host "  [+] $_" -ForegroundColor Magenta }
                Write-Host "  ... and $($extraFolders.Count - 20) more." -ForegroundColor Magenta
            } else {
                $extraFolders | ForEach-Object { Write-Host "  [+] $_" -ForegroundColor Magenta }
            }
            
            if ($dataOutDir) {
                $extraLogPath = Join-Path $dataOutDir "$($mailboxSafe)_ExtraFolders_$timestamp.txt"
                $extraFolders | Out-File -FilePath $extraLogPath -Encoding UTF8
                Write-LogMessage -Level WARN -Message "Extra folders exported to: $extraLogPath"
                Write-Host "Saved extra list: $extraLogPath" -ForegroundColor Magenta
            }
        }

    } else {
         Write-Error "Expected structure file not found at: $ExpectedStructureFile"
         Write-LogMessage -Level ERROR -Message "Expected structure file not found at: $ExpectedStructureFile"
    }
}
