<#
.SYNOPSIS
Utility functions for logging and managing log files.

.DESCRIPTION
This module provides functions to:
- Set a log file path.
- Clear the log file path.
- Write log messages with different levels to a log file.

.EXPORTED FUNCTIONS
- Write-LogMessage
- Set-LogFile
- Clear-LogFile
#>

# Internal variable to hold the optional log file path
$script:LogFilePath = $null

<#
.SYNOPSIS
Sets the log file path.

.PARAMETER Path
The full path to the log file.

.DESCRIPTION
This function sets the path for the log file and ensures the directory exists.
#>
function Set-LogFile {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Path
    )
    if ($PSCmdlet.ShouldProcess($Path, 'Set log file')) {
        $script:LogFilePath = $Path
        # Ensure directory exists
        $dir = Split-Path -Path $script:LogFilePath -Parent
        if (-not (Test-Path $dir)) {
            New-Item -Path $dir -ItemType Directory -Force | Out-Null
        }
    }
}

<#
.SYNOPSIS
Clears the log file path.

.DESCRIPTION
This function clears the internal log file path, effectively disabling logging to a file.
#>
function Clear-LogFile {
    [CmdletBinding(SupportsShouldProcess=$true)]
    param()
    if ($PSCmdlet.ShouldProcess('LogFilePath', 'Clear log file')) {
        $script:LogFilePath = $null
    }
}

<#
.SYNOPSIS
Writes a log message to the log file.

.PARAMETER Message
The log message to write.

.PARAMETER Level
The log level (INFO, WARN, ERROR, DEBUG, SUCCESS). Default is INFO.

.PARAMETER FilePath
Optional. The path to the log file. Overrides the default or set log file path.

.PARAMETER NoConsole
If specified, suppresses console output.

.DESCRIPTION
This function writes a log message to the specified log file or a default log file.
If the log file path is not set, it creates a default log file in the `logs` directory.
#>
function Write-LogMessage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [ValidateSet('INFO','WARN','ERROR','DEBUG','SUCCESS')]
        [string]$Level = 'INFO',
        
        [Parameter(Mandatory=$false)]
        [string]$FilePath,

        [Parameter(Mandatory=$false)]
        [switch]$NoConsole
    )
    $ts = (Get-Date).ToString('s')
    $entry = "[$ts] [$Level] $Message"
    
    # Console output with color
    if (-not $NoConsole) {
        $color = switch ($Level) {
            'INFO'    { 'White' }
            'SUCCESS' { 'Green' }
            'WARN'    { 'Yellow' }
            'ERROR'   { 'Red' }
            'DEBUG'   { 'Gray' }
            default   { 'White' }
        }
        Write-Host $entry -ForegroundColor $color
    }

    # Determine the log file path
    $target = $FilePath
    if (-not $target -and $script:LogFilePath) {
        $target = $script:LogFilePath
    }
    if (-not $target) {
        $defaultDir = Join-Path -Path (Get-Location) -ChildPath 'logs'
        if (-not (Test-Path $defaultDir)) {
            New-Item -Path $defaultDir -ItemType Directory -Force | Out-Null
        }
        $target = Join-Path -Path $defaultDir -ChildPath 'script.log'
    }

    # Write the log entry to the file
    try {
        $entry | Out-File -FilePath $target -Append -Encoding utf8
    }
    catch {
        Write-Warning "Failed to write to log file '$target': $($_.Exception.Message)"
    }
}

<#
.SYNOPSIS
Writes a progress message with optional logging.

.PARAMETER Activity
The activity name.

.PARAMETER Status
The status message.

.PARAMETER PercentComplete
Optional percentage complete (0-100).

.PARAMETER Id
Progress bar ID for nested progress bars.

.EXAMPLE
Write-ProgressUpdate -Activity "Processing files" -Status "File 5 of 10" -PercentComplete 50
#>
function Write-ProgressUpdate {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Activity,
        
        [Parameter(Mandatory=$false)]
        [string]$Status = '',

        [Parameter(Mandatory=$false)]
        [ValidateRange(0,100)]
        [int]$PercentComplete = -1,

        [Parameter(Mandatory=$false)]
        [int]$Id = 0,

        [Parameter(Mandatory=$false)]
        [switch]$Completed
    )

    $params = @{
        Activity = $Activity
        Id = $Id
    }
    
    if ($Completed) {
        $params['Completed'] = $true
    }
    else {
        if ($Status) { $params['Status'] = $Status }
        if ($PercentComplete -ge 0) { $params['PercentComplete'] = $PercentComplete }
    }

    Write-Progress @params
    
    # Log significant progress milestones
    if ($PercentComplete -in @(0, 25, 50, 75, 100) -or $Completed) {
        $logMsg = if ($Completed) { "$Activity - Completed" } else { "$Activity - $Status ($PercentComplete%)" }
        Write-LogMessage -Message $logMsg -Level 'DEBUG' -NoConsole
    }
}

# Export only the intended functions
if ($null -ne $PSModuleInfo) {
    Export-ModuleMember -Function Write-LogMessage, Set-LogFile, Clear-LogFile, Write-ProgressUpdate
}
