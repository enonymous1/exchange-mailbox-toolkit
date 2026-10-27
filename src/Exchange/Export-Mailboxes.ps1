<#
.SYNOPSIS
    Orchestrate exporting across many mailboxes using Get-Mailboxes + Export-SharedCalendar
.DESCRIPTION
    Retrieves a mailbox list (Shared/User/All, internal/external) and invokes
    `Export-SharedCalendar.ps1` for each mailbox. Runs each mailbox export in a separate
    PowerShell process to isolate failures.
#>
#requires -Version 7.2

[CmdletBinding()]
param(
    [ValidateSet('All','Shared','User')]
    [string]$MailboxType = 'Shared',

    [switch]$InternalOnly,
    [switch]$ExternalOnly,
    [string[]]$Domains,

    [Parameter(Mandatory=$false)]
    [int]$Max = 0,

    [Parameter(Mandatory=$false)]
    [string]$OutDir,

    [Parameter(Mandatory=$false)]
    [switch]$WhatIf,

    [Parameter(Mandatory=$false)]
    [switch]$ConfirmRun
)

# Dot-source helpers
. "$PSScriptRoot/Get-GraphMailboxes.ps1"

# Ensure Graph connection exists for mailbox discovery
$connectHelper = Join-Path $PSScriptRoot '../Utils/Connect-ExchangeEnvironment.ps1'
if (Test-Path $connectHelper) {
    try { . $connectHelper } catch { Write-Warning "Could not dot-source Connect-ExchangeEnvironment: $($_.Exception.Message)" }
}

if (Get-Command -Name Connect-ExchangeEnvironment -ErrorAction SilentlyContinue) {
    try {
        Write-Verbose "Establishing Microsoft Graph connection for orchestrator..."
        $status = Connect-ExchangeEnvironment -GraphScopes @('User.Read.All','Directory.Read.All') -LogFile $null -Verbose
    }
    catch {
        Write-Error "Connect-ExchangeEnvironment failed: $($_.Exception.Message)"
    }
    if (-not $status -or -not $status.Graph.Connected) {
        Write-Error "Could not establish Microsoft Graph connection for mailbox discovery. Run 'Connect-ExchangeEnvironment' interactively or supply credentials suitable for non-interactive authentication."
        exit 1
    }
}
else {
    Write-Warning "Connect-ExchangeEnvironment helper not available; Get-Mailboxes will attempt to authenticate itself."
}

Write-Verbose "Orchestration: fetching mailboxes (Type=$MailboxType, InternalOnly=$InternalOnly, ExternalOnly=$ExternalOnly, Max=$Max)"
$mbs = Get-GraphMailboxes -Type $MailboxType -InternalOnly:$InternalOnly -ExternalOnly:$ExternalOnly -Domains $Domains -Max $Max
Write-Verbose "Found $($mbs.Count) mailboxes (Max=$Max)"

if ($mbs.Count -eq 0) { Write-Warning 'No mailboxes found; aborting.'; exit 0 }
if ($ConfirmRun -and -not $WhatIf) {
    $ok = Read-Host "Proceed to export $($mbs.Count) mailboxes? (Y/n)"
    if ($ok -and $ok.Trim().ToLower() -eq 'n') { Write-Output 'Aborting by user request.'; exit 0 }
}

$results = [System.Collections.Generic.List[PSObject]]::new()
$scriptPath = Join-Path $PSScriptRoot 'Export-SharedCalendar.ps1'

foreach ($m in $mbs) {
    $args = @('-NoLogo','-NoProfile','-File',$scriptPath,'-SharedMailboxEmail',$m.Mail)
    if ($OutDir) { $args += @('-OutDir',$OutDir) }
    if ($WhatIf) { $args += '-WhatIf' }
    $argString = $args -join ' '
    Write-Verbose "Invoking export for $($m.Mail)"
    $proc = Start-Process -FilePath pwsh -ArgumentList $args -NoNewWindow -Wait -PassThru
    $ok = $proc.ExitCode -eq 0
    $results.Add([pscustomobject]@{ Mail=$m.Mail; DisplayName=$m.DisplayName; ExitCode=$proc.ExitCode; Success=$ok })
}

$summary = $results | Group-Object -Property Success | ForEach-Object { [pscustomobject]@{ Success=$_.Name; Count=$_.Count } }
Write-Output "Finished. Results:"
$results | Format-Table -AutoSize
$summary | Format-Table -AutoSize

# Optionally write a summary file
$logDir = Join-Path (Resolve-Path "$PSScriptRoot/../..").Path 'logs'
if (-not (Test-Path $logDir)) { New-Item -Path $logDir -ItemType Directory -Force | Out-Null }
$timestamp = (Get-Date).ToString('yyyyMMddHHmmss')
$summaryFile = Join-Path $logDir ("export_mailboxes_summary_{0}.json" -f $timestamp)
$results | ConvertTo-Json -Depth 5 | Set-Content -Path $summaryFile -Force
Write-Output "Orchestration summary written to: $summaryFile"