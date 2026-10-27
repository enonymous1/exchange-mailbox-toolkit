<#
.SYNOPSIS
Centralized Authentication & Environment Loader for Exchange Online / Microsoft 365 administration.

.DESCRIPTION
Handles connections to Microsoft Graph and Exchange Online.
Checks existing sessions to avoid unnecessary re-authentication prompts.

.PARAMETER GraphScopes
List of Graph API permissions required by the calling script.

.PARAMETER LoadExchange
Switch to connect to Exchange Online Management (EXO).

.PARAMETER ForceReauth
Forces re-authentication even if a session already exists.

.PARAMETER LogFile
Optional. Path to a log file for recording connection details.

.EXAMPLE
. "$PSScriptRoot/Connect-ExchangeEnvironment.ps1" -GraphScopes "User.Read.All", "Calendars.Read" -LoadExchange
#>

function Connect-ExchangeEnvironment {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false)]
        [string[]]$GraphScopes = @("User.Read.All"),

        [Parameter(Mandatory=$false)]
        [switch]$LoadExchange,

        [Parameter(Mandatory=$false)]
        [switch]$ForceReauth,

        [Parameter(Mandatory=$false)]
        [string]$LogFile = ''
    )

    # NOTE (v2.0.0): the old Helpers.ps1 reference was removed. Logging helpers
    # (Write-LogMessage) are loaded by the caller or by the ExchangeMailboxToolkit module;
    # Connect-ExchangeEnvironment already falls back to Write-Verbose when they are absent.

    # Initialize status object
    $status = [PSCustomObject]@{
        Graph = [PSCustomObject]@{ Connected = $false; Account = $null; Scopes = @(); Message = $null }
        Exchange = [PSCustomObject]@{ Connected = $false; Message = $null }
        Modules = [PSCustomObject]@{ MicrosoftGraph = $false; ExchangeOnlineManagement = $false }
        Warnings = @()
        Errors = @()
    }

    Write-Verbose "--- INITIALIZING EXCHANGE ENVIRONMENT ---"

    # Helper function for logging
    function Write-Info {
        param([string]$Message)
        if (Get-Command -Name Write-LogMessage -ErrorAction SilentlyContinue) {
            Write-LogMessage -Message $Message -Level 'INFO'
        } else {
            Write-Verbose $Message
        }
        if ($LogFile) {
            $Message | Out-File -FilePath $LogFile -Append -Encoding UTF8
        }
    }

    # Check module availability
    $status.Modules.MicrosoftGraph = $null -ne (Get-Command -Name Connect-MgGraph -ErrorAction SilentlyContinue)
    $status.Modules.ExchangeOnlineManagement = $null -ne (Get-Command -Name Connect-ExchangeOnline -ErrorAction SilentlyContinue)

    if (-not $status.Modules.MicrosoftGraph -and $GraphScopes) {
        $msg = 'Microsoft Graph PowerShell module is not available. Install Microsoft.Graph module.'
        Write-Warning $msg
        $status.Warnings += $msg
        $status.Graph.Message = $msg
    }

    if (-not $status.Modules.ExchangeOnlineManagement -and $LoadExchange) {
        $msg = 'ExchangeOnlineManagement module is not available. Install ExchangeOnlineManagement module.'
        Write-Warning $msg
        $status.Warnings += $msg
        $status.Exchange.Message = $msg
    }

    # Connect to Microsoft Graph
    if ($GraphScopes -and $status.Modules.MicrosoftGraph) {
        Write-Info "Checking Microsoft Graph connection..."
        try {
            $current = Get-MgContext -ErrorAction SilentlyContinue
            $currentScopes = $current?.Scopes ?? @()
            $missingScopes = $GraphScopes | Where-Object { -not ($currentScopes -contains $_) }

            if ($ForceReauth -or -not $current -or $missingScopes) {
                Write-Info "Authenticating to Microsoft Graph..."
                Connect-MgGraph -Scopes $GraphScopes -NoWelcome -ErrorAction Stop
                $ctx = Get-MgContext -ErrorAction Stop
                $status.Graph.Connected = $true
                $status.Graph.Account = $ctx.Account
                $status.Graph.Scopes = $ctx.Scopes
            } else {
                $status.Graph.Connected = $true
                $status.Graph.Account = $current.Account
                $status.Graph.Scopes = $currentScopes
            }
        } catch {
            $err = "Failed to connect to Microsoft Graph: $($_.Exception.Message)"
            Write-Warning $err
            $status.Errors += $err
            $status.Graph.Message = $err
        }
    }

    # Connect to Exchange Online
    if ($LoadExchange -and $status.Modules.ExchangeOnlineManagement) {
        Write-Info "Checking Exchange Online connection..."
        try {
            Connect-ExchangeOnline -ShowProgress:$false -ErrorAction Stop
            $status.Exchange.Connected = $true
        } catch {
            $err = "Failed to connect to Exchange Online: $($_.Exception.Message)"
            Write-Warning $err
            $status.Errors += $err
            $status.Exchange.Message = $err
        }
    }

    Write-Verbose "--- ENVIRONMENT READY ---"
    return $status
}

if ($null -ne $PSModuleInfo) {
    Export-ModuleMember -Function Connect-ExchangeEnvironment
}

# NOTE (v2.0.0): importing this file used to run Connect-ExchangeEnvironment as a side
# effect. It no longer does: call Connect-ExchangeEnvironment explicitly after importing,
# so importing the module or dot-sourcing the helpers never triggers an
# interactive sign-in on its own.