<#
.SYNOPSIS
    Check whether the signed-in account can read a user through Microsoft Graph.

.DESCRIPTION
    v2.1.0: folded in from mp-powershell ("User Impersonation.ps1"), with
    its central claim corrected. The original was titled impersonation,
    but it never impersonated anyone: Connect-MgGraph with User.Read.All
    followed by Get-MgUser only proves that the signed-in account can
    read that user. It does not act as the user, and no mailbox, message,
    or send-as action is performed. This script is renamed to say what
    it actually does, and it returns a result object instead of
    Write-Host / Format-List output so it can be piped and tested.

    Changes from the original, besides the rename:
    - No unused Connect-ExchangeOnline: the original opened an Exchange
      session it never used and never disconnected. This script only
      connects to Graph, through Connect-ExchangeEnvironment.
    - Least privilege: only User.Read.All is requested. The original
      also requested Group.Read.All and never used it.
    - A -UserPrincipalName parameter replaces the hard-coded placeholder
      address, and the script no longer prints "Impersonation
      successful!".
    - Failure is a result object (CanRead = $false with the error), not
      a script crash: a failed check is the answer, not an exception.

    This is read-only and changes nothing in the tenant. It is NOT an
    impersonation test. True application impersonation in Exchange
    Online is governed by RBAC for Applications / Application Access
    Policies, which this script does not exercise.

.PARAMETER UserPrincipalName
    The user to read, for example user@contoso.com. Required when run
    as a script; the function parameter is mandatory.

.EXAMPLE
    pwsh -File src/Exchange/Test-GraphUserAccess.ps1 -UserPrincipalName 'user@contoso.com'
#>
#requires -Version 7.2

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$UserPrincipalName
)

. (Join-Path $PSScriptRoot '..' 'Utils' 'Initialize-Environment.ps1')

function Test-GraphUserAccess {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserPrincipalName
    )

    $envStatus = Connect-ExchangeEnvironment -GraphScopes @('User.Read.All') -ErrorAction Stop
    if (-not $envStatus.Graph.Connected) {
        throw "Microsoft Graph is not connected: $($envStatus.Graph.Message) Install the Microsoft.Graph module and sign in, then retry."
    }

    if (-not (Get-Command -Name Get-MgUser -ErrorAction SilentlyContinue)) {
        throw 'Get-MgUser is not available. Install the Microsoft.Graph module and reconnect.'
    }

    try {
        $user = Get-MgUser -UserId $UserPrincipalName -ErrorAction Stop
    } catch {
        Write-Warning "Could not read user ${UserPrincipalName}: $($_.Exception.Message)"
        return [PSCustomObject]@{
            UserPrincipalName = $UserPrincipalName
            DisplayName       = $null
            CanRead           = $false
            Error             = $_.Exception.Message
        }
    }

    if ($null -eq $user) {
        return [PSCustomObject]@{
            UserPrincipalName = $UserPrincipalName
            DisplayName       = $null
            CanRead           = $false
            Error             = 'User not found or no data returned.'
        }
    }

    return [PSCustomObject]@{
        UserPrincipalName = $UserPrincipalName
        DisplayName       = $user.DisplayName
        CanRead           = $true
        Error             = $null
    }
}

if ($MyInvocation.InvocationName -ne '.') {
    if (-not $UserPrincipalName) {
        throw 'Test-GraphUserAccess requires -UserPrincipalName, for example user@contoso.com.'
    }
    Test-GraphUserAccess -UserPrincipalName $UserPrincipalName
}
