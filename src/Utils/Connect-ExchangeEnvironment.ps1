<#
.SYNOPSIS
    Loader for Connect-ExchangeEnvironment.

.DESCRIPTION
    The Connect-ExchangeEnvironment function is defined in Initialize-Environment.ps1 in
    this folder. Older scripts, the README, and the Copilot instructions
    reference this loader file (under its former name in v1, which did not exist at all),
    so those callers failed to load the function. This file exists
    so those references work: it simply loads Initialize-Environment.ps1.

    New code should import the ExchangeMailboxToolkit module or dot-source
    Initialize-Environment.ps1 directly. Importing this file defines the
    function only; it never connects on its own.
#>

. (Join-Path $PSScriptRoot 'Initialize-Environment.ps1')
