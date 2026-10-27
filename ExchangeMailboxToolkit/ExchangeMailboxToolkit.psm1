<#
.SYNOPSIS
    ExchangeMailboxToolkit module loader.

.DESCRIPTION
    Loads the shared helpers that used to be dot-sourced ad hoc by every
    script: logging, Graph helpers, the Connect-ExchangeEnvironment environment loader,
    and the mailbox discovery functions. Importing the module defines the
    functions only; it never signs in to Graph or Exchange on its own
    (that side effect was removed from Initialize-Environment.ps1 in
    v2.0.0).

    The Exchange scripts under src/ still work standalone by dot-sourcing
    the same files, so this module is an addition, not a breaking change
    to how those scripts are invoked.

.EXAMPLE
    Import-Module ./ExchangeMailboxToolkit
    Connect-ExchangeEnvironment -GraphScopes 'User.Read.All','Mail.Read.Shared'
#>

$moduleRoot = $PSScriptRoot
$utilsRoot = Join-Path $moduleRoot '..' 'src' 'Utils'
$exchangeRoot = Join-Path $moduleRoot '..' 'src' 'Exchange'

$filesToLoad = @(
    (Join-Path $utilsRoot 'LoggingHelpers.ps1')
    (Join-Path $utilsRoot 'GraphHelpers.ps1')
    (Join-Path $utilsRoot 'Initialize-Environment.ps1')
    (Join-Path $exchangeRoot 'Get-GraphMailboxes.ps1')
    (Join-Path $exchangeRoot 'Get-Mailboxes.ps1')
)

foreach ($file in $filesToLoad) {
    if (Test-Path -Path $file) {
        . $file
    } else {
        Write-Warning "ExchangeMailboxToolkit: expected helper file not found: $file"
    }
}

Export-ModuleMember -Function @(
    'Connect-ExchangeEnvironment'
    'Write-LogMessage', 'Set-LogFile', 'Clear-LogFile', 'Write-ProgressUpdate'
    'Get-SafeFilename', 'Get-EncodedGraphId', 'Invoke-GraphWithRetry'
    'Move-GraphMessage', 'Get-MailFoldersWithPaths', 'Invoke-EnsureFolderPath'
    'Get-GraphMailboxes', 'Get-Mailboxes'
)
