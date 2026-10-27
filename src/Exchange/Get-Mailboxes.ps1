<# Compatibility shim for older callers. Calls into the new Get-GraphMailboxes function.
   Keep this wrapper to avoid breaking scripts that still call Get-Mailboxes.
#>
#requires -Version 7.2

if (Test-Path (Join-Path $PSScriptRoot 'Get-GraphMailboxes.ps1')) {
    . "$PSScriptRoot/Get-GraphMailboxes.ps1"
}

function Get-Mailboxes {
    [CmdletBinding()]
    param(
        [ValidateSet('All','Shared','User')]
        [string]$Type = 'Shared',

        [Parameter()]
        [switch]$InternalOnly,

        [Parameter()]
        [switch]$ExternalOnly,

        [Parameter()]
        [string[]]$Domains,

        [Parameter()]
        [int]$Max = 0
    )

    Write-Warning "Get-Mailboxes is deprecated; use Get-GraphMailboxes instead. Forwarding call."
    return Get-GraphMailboxes -Type $Type -InternalOnly:$InternalOnly -ExternalOnly:$ExternalOnly -Domains $Domains -Max $Max
}

try {
    Export-ModuleMember -Function Get-Mailboxes -ErrorAction SilentlyContinue
}
catch {
    # Not running as a module - ignore
}
