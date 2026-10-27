<#
.SYNOPSIS
Helper functions for Microsoft Graph API operations.

.DESCRIPTION
This module provides reusable functions for:
- Safe filename generation
- URL encoding for Graph API calls
- Retry logic with exponential backoff
- Message moving operations

.EXPORTED FUNCTIONS
- Get-SafeFilename
- Invoke-GraphWithRetry
- Move-GraphMessage
- Get-MailFoldersWithPaths
#>

<#
.SYNOPSIS
Converts a string to a safe filename by replacing invalid characters.

.PARAMETER Name
The string to convert to a safe filename.

.EXAMPLE
Get-SafeFilename -Name "user@domain.com"
# Returns: user_domain_com
#>
function Get-SafeFilename {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false)]
        [string]$Name
    )
    if ($null -eq $Name -or [string]::IsNullOrWhiteSpace($Name)) { 
        return 'unnamed' 
    }
    $safe = $Name -replace '[\\/:*?"<>|]', '_'
    $safe = $safe -replace "[\r\n\t]", ' '
    return $safe.Trim()
}

<#
.SYNOPSIS
URL-encodes a string for use in Graph API URIs.

.PARAMETER Value
The string to encode.

.EXAMPLE
Get-EncodedGraphId -Value "AAMkADU5..."
#>
function Get-EncodedGraphId {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Value
    )
    return [System.Uri]::EscapeDataString($Value)
}

<#
.SYNOPSIS
Invokes a Graph API request with automatic retry logic and exponential backoff.

.PARAMETER Method
HTTP method (GET, POST, DELETE, PATCH, PUT).

.PARAMETER Uri
The Graph API URI to call.

.PARAMETER Body
Optional request body (for POST/PATCH/PUT).

.PARAMETER MaxRetries
Maximum number of retry attempts. Default is 3.

.PARAMETER InitialDelaySeconds
Initial delay in seconds before first retry. Default is 2.

.EXAMPLE
Invoke-GraphWithRetry -Method GET -Uri "users/user@domain.com"
#>
function Invoke-GraphWithRetry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateSet('GET','POST','DELETE','PATCH','PUT')]
        [string]$Method,

        [Parameter(Mandatory=$true)]
        [string]$Uri,

        [Parameter(Mandatory=$false)]
        [string]$Body,

        [Parameter(Mandatory=$false)]
        [int]$MaxRetries = 3,

        [Parameter(Mandatory=$false)]
        [int]$InitialDelaySeconds = 2
    )

    $attempt = 0
    $delay = $InitialDelaySeconds
    $lastError = $null

    while ($attempt -lt $MaxRetries) {
        try {
            $params = @{
                Method = $Method
                Uri = $Uri
                ErrorAction = 'Stop'
            }
            if ($Body) {
                $params['Body'] = $Body
                $params['ContentType'] = 'application/json'
            }
            
            $result = Invoke-MgGraphRequest @params
            return $result
        }
        catch {
            $lastError = $_
            $attempt++
            
            # Check if it's a throttling error (429) or transient error (5xx)
            $statusCode = $_.Exception.Response.StatusCode.value__
            $isRetryable = $statusCode -eq 429 -or ($statusCode -ge 500 -and $statusCode -lt 600)
            
            if ($attempt -lt $MaxRetries -and $isRetryable) {
                Write-Verbose "Request failed (attempt $attempt/$MaxRetries). Retrying in $delay seconds..."
                Start-Sleep -Seconds $delay
                $delay = $delay * 2  # Exponential backoff
            }
            else {
                break
            }
        }
    }
    
    # All retries exhausted
    throw $lastError
}

<#
.SYNOPSIS
Moves a message to a destination folder with fallback strategies and retry logic.

.PARAMETER UserId
The user ID or UPN.

.PARAMETER MessageId
The message ID to move.

.PARAMETER SourceFolderId
The source folder ID.

.PARAMETER DestinationFolderId
The destination folder ID.

.PARAMETER MaxAttempts
Maximum number of different strategies to attempt. Default is 3.

.EXAMPLE
Move-GraphMessage -UserId "user@domain.com" -MessageId "AAMk..." -SourceFolderId "inbox" -DestinationFolderId "folder123"
#>
function Move-GraphMessage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$UserId,

        [Parameter(Mandatory=$true)]
        [string]$MessageId,

        [Parameter(Mandatory=$true)]
        [string]$SourceFolderId,

        [Parameter(Mandatory=$true)]
        [string]$DestinationFolderId,

        [Parameter(Mandatory=$false)]
        [int]$MaxAttempts = 3
    )

    $encMid = Get-EncodedGraphId -Value $MessageId
    $encFolder = Get-EncodedGraphId -Value $SourceFolderId
    $body = @{ destinationId = $DestinationFolderId } | ConvertTo-Json
    
    $strategies = @(
        @{ Name = 'Primary'; Uri = "users/$UserId/messages/$encMid/move" },
        @{ Name = 'FolderScoped'; Uri = "users/$UserId/mailFolders/$encFolder/messages/$encMid/move" },
        @{ Name = 'DirectAPI'; Uri = "users/$UserId/mailFolders/$encFolder/messages/$encMid/move" }
    )

    $triedUris = [System.Collections.ArrayList]::new()
    $lastError = $null

    foreach ($strategy in $strategies[0..([Math]::Min($MaxAttempts-1, $strategies.Count-1))]) {
        try {
            Write-Verbose "Attempting to move message using strategy: $($strategy.Name)"
            [void]$triedUris.Add($strategy.Uri)
            
            $result = Invoke-GraphWithRetry -Method POST -Uri $strategy.Uri -Body $body
            
            return @{
                Success = $true
                Strategy = $strategy.Name
                TriedUris = $triedUris -join '; '
                Result = $result
            }
        }
        catch {
            $lastError = $_
            Write-Verbose "Strategy $($strategy.Name) failed: $($_.Exception.Message)"
            continue
        }
    }

    # All strategies failed
    return @{
        Success = $false
        Strategy = $null
        TriedUris = $triedUris -join '; '
        Error = $lastError.Exception.Message
        ErrorDetails = $lastError
    }
}

<#
.SYNOPSIS
Recursively enumerates all mail folders for a user and builds full paths.

.PARAMETER UserId
The user ID or UPN.

.PARAMETER ProgressActivity
Optional activity name for progress bar.

.EXAMPLE
$folders = Get-MailFoldersWithPaths -UserId "user@domain.com"
#>
function Get-MailFoldersWithPaths {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$UserId,

        [Parameter(Mandatory=$false)]
        [string]$ProgressActivity = 'Enumerating folders'
    )

    $list = [System.Collections.ArrayList]::new()
    $processed = 0

    function Recurse-Folders {
        param(
            [string]$folderId,
            [string]$parentPath
        )
        
        try {
            # Use direct REST API to avoid SDK caching issues
            $encFolderId = Get-EncodedGraphId -Value $folderId
            $uri = "https://graph.microsoft.com/v1.0/users/$UserId/mailFolders/$encFolderId/childFolders"
            $response = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
            
            $children = $response.value
            if ($response.'@odata.nextLink') {
                # Handle pagination
                $nextLink = $response.'@odata.nextLink'
                while ($nextLink) {
                    $response = Invoke-MgGraphRequest -Method GET -Uri $nextLink -ErrorAction Stop
                    $children += $response.value
                    $nextLink = $response.'@odata.nextLink'
                }
            }
            
            foreach ($c in $children) {
                $path = if ($parentPath) { "$parentPath\$($c.displayName)" } else { $c.displayName }
                [void]$list.Add([PSCustomObject]@{
                    Id = $c.id
                    DisplayName = $c.displayName
                    Path = $path
                    ParentFolderId = $folderId
                    ChildFolderCount = $c.childFolderCount
                    TotalItemCount = $c.totalItemCount
                    UnreadItemCount = $c.unreadItemCount
                })
                $script:processed++
                if ($script:processed % 10 -eq 0) {
                    Write-Progress -Activity $ProgressActivity -Status "Found $script:processed folders" -Id 1
                }
                Recurse-Folders -folderId $c.id -parentPath $path
            }
        }
        catch {
            Write-Warning "Failed to enumerate child folders for $folderId : $($_.Exception.Message)"
        }
    }

    # Get root-level folders using direct REST API
    try {
        $uri = "https://graph.microsoft.com/v1.0/users/$UserId/mailFolders"
        $response = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
        
        $roots = $response.value
        if ($response.'@odata.nextLink') {
            # Handle pagination
            $nextLink = $response.'@odata.nextLink'
            while ($nextLink) {
                $response = Invoke-MgGraphRequest -Method GET -Uri $nextLink -ErrorAction Stop
                $roots += $response.value
                $nextLink = $response.'@odata.nextLink'
            }
        }
        
        foreach ($r in $roots) {
            $path = $r.displayName
            [void]$list.Add([PSCustomObject]@{
                Id = $r.id
                DisplayName = $r.displayName
                Path = $path
                ParentFolderId = $null
                ChildFolderCount = $r.childFolderCount
                TotalItemCount = $r.totalItemCount
                UnreadItemCount = $r.unreadItemCount
            })
            $processed++
            Recurse-Folders -folderId $r.id -parentPath $path
        }
    }
    catch {
        Write-Error "Failed to get root folders: $($_.Exception.Message)"
        return $list
    }
    finally {
        Write-Progress -Activity $ProgressActivity -Completed -Id 1
    }

    Write-Verbose "Enumerated $($list.Count) total folders"
    return $list
}

<#
.SYNOPSIS
Gets a fresh folder object by navigating its path. Returns folder with current valid ID.

.DESCRIPTION
Queries folders on-demand to avoid stale cached IDs. Navigates from root following the path.

.PARAMETER UserId
The user ID or UPN.

.PARAMETER FolderPath
The full folder path (e.g., "Inbox\Projects\Qualified").

.EXAMPLE
$folder = Get-FreshFolderByPath -UserId "user@domain.com" -FolderPath "Inbox\Archive\2025"
#>
function Get-FreshFolderByPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$UserId,

        [Parameter(Mandatory=$true)]
        [string]$FolderPath
    )

    $parts = $FolderPath -split '\\'
    $currentFolder = $null

    for ($i = 0; $i -lt $parts.Count; $i++) {
        $folderName = $parts[$i]
        
        if ($i -eq 0) {
            # Root level - query all root folders
            try {
                $uri = "https://graph.microsoft.com/v1.0/users/$UserId/mailFolders"
                $response = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
                $currentFolder = $response.value | Where-Object { $_.displayName -eq $folderName } | Select-Object -First 1
            }
            catch {
                Write-Error "Failed to query root folders: $($_.Exception.Message)"
                return $null
            }
        }
        else {
            # Child folder - query children of current folder
            try {
                $uri = "https://graph.microsoft.com/v1.0/users/$UserId/mailFolders/$($currentFolder.id)/childFolders"
                $response = Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
                $currentFolder = $response.value | Where-Object { $_.displayName -eq $folderName } | Select-Object -First 1
            }
            catch {
                Write-Error "Failed to query child folders of '$($parts[0..$($i-1)] -join '\')': $($_.Exception.Message)"
                return $null
            }
        }

        if (-not $currentFolder) {
            Write-Verbose "Folder not found: $FolderPath (stopped at part $($i): '$folderName')"
            return $null
        }
    }

    return $currentFolder
}

<#
.SYNOPSIS
Gets or creates a folder by path, creating missing parent folders as needed.

.PARAMETER UserId
The user ID or UPN.

.PARAMETER FolderPath
The full folder path (e.g., "Inbox\Projects\2026").

.PARAMETER FolderMap
Hashtable mapping lowercase paths to folder objects.

.PARAMETER DryRun
If specified, simulates folder creation without making actual API calls.

.EXAMPLE
$result = Invoke-EnsureFolderPath -UserId "user@domain.com" -FolderPath "Inbox\Archive\2026" -FolderMap $map
#>
function Invoke-EnsureFolderPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$UserId,

        [Parameter(Mandatory=$true)]
        [string]$FolderPath,

        [Parameter(Mandatory=$true)]
        [hashtable]$FolderMap,

        [Parameter(Mandatory=$false)]
        [switch]$DryRun
    )

    $parts = $FolderPath -split '\\'
    $currentPath = ''
    $parentId = $null
    $created = [System.Collections.ArrayList]::new()

    foreach ($p in $parts) {
        $currentPath = if ($currentPath) { "$currentPath\$p" } else { $p }
        
        if ($FolderMap.ContainsKey($currentPath.ToLower())) {
            $parentId = $FolderMap[$currentPath.ToLower()].Id
            continue
        }

        # Need to create this folder
        if ($DryRun) {
            Write-Verbose "(DryRun) Would create folder: $currentPath"
            $fakeId = [guid]::NewGuid().ToString()
            $FolderMap[$currentPath.ToLower()] = [PSCustomObject]@{
                Id = $fakeId
                DisplayName = $p
                Path = $currentPath
            }
            $parentId = $fakeId
            [void]$created.Add($currentPath)
        }
        else {
            Write-Verbose "Creating folder: $currentPath"
            try {
                $body = @{ displayName = $p } | ConvertTo-Json
                $uri = if ($parentId) {
                    "users/$UserId/mailFolders/$parentId/childFolders"
                } else {
                    "users/$UserId/mailFolders"
                }
                
                $resp = Invoke-GraphWithRetry -Method POST -Uri $uri -Body $body
                
                # Add to map
                $FolderMap[$currentPath.ToLower()] = [PSCustomObject]@{
                    Id = $resp.id
                    DisplayName = $resp.displayName
                    Path = $currentPath
                }
                $parentId = $resp.id
                [void]$created.Add($currentPath)
                Write-Verbose "Created: $currentPath (Id: $parentId)"
            }
            catch {
                Write-Warning "Failed to create folder ${currentPath}: $($_.Exception.Message)"
                return @{
                    Success = $false
                    Created = $created
                    Error = $_.Exception.Message
                }
            }
        }
    }

    return @{
        Success = $true
        Created = $created
        FolderId = $parentId
    }
}

# Export module members
if ($null -ne $PSModuleInfo) {
    Export-ModuleMember -Function Get-SafeFilename, Get-EncodedGraphId, Invoke-GraphWithRetry, `
        Move-GraphMessage, Get-MailFoldersWithPaths, Invoke-EnsureFolderPath
}
