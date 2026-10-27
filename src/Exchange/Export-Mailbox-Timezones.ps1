<#
.SYNOPSIS
    Export mailbox time zone and regional settings to CSV.

.DESCRIPTION
    v2.0.0: this file was a placeholder in v1 that printed a message and
    did nothing. It now does the job the TODO described: it connects to
    Exchange Online via Connect-ExchangeEnvironment, enumerates mailboxes, and exports
    each mailbox's regional configuration (time zone, language, date and
    time formats) to CSV.

    This needs a real Exchange Online session, so it cannot be verified
    on a machine without EXO access. Run it scoped to one mailbox first:
    it is read-only, but it does call Get-EXOMailbox across the tenant
    when no mailbox is given.

.PARAMETER MailboxEmail
    Optional. Export only this mailbox instead of the whole tenant.

.PARAMETER OutputPath
    Path for the CSV. Defaults to data/output/mailbox_timezones.csv
    under the current directory.

.EXAMPLE
    pwsh -File src/Exchange/Export-Mailbox-Timezones.ps1 -MailboxEmail user@contoso.com
#>
#requires -Version 7.2

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [string]$MailboxEmail,

    [Parameter(Mandatory=$false)]
    [string]$OutputPath = "data/output/mailbox_timezones.csv"
)

. (Join-Path $PSScriptRoot '..' 'Utils' 'Initialize-Environment.ps1')

$envStatus = Connect-ExchangeEnvironment -GraphScopes @('User.Read.All') -LoadExchange -ErrorAction Stop
if (-not $envStatus.Exchange.Connected) {
    throw "Exchange Online is not connected: $($envStatus.Exchange.Message) Install the ExchangeOnlineManagement module and sign in, then retry."
}

if (-not (Get-Command -Name Get-EXOMailbox -ErrorAction SilentlyContinue)) {
    throw 'Get-EXOMailbox is not available. Install the ExchangeOnlineManagement module and reconnect with -LoadExchange.'
}

if ($MailboxEmail) {
    $mailboxes = @(Get-EXOMailbox -Identity $MailboxEmail -ErrorAction Stop)
} else {
    $mailboxes = @(Get-EXOMailbox -ResultSize Unlimited -ErrorAction Stop)
}

$rows = foreach ($mb in $mailboxes) {
    try {
        $regional = Get-MailboxRegionalConfiguration -Identity $mb.UserPrincipalName -ErrorAction Stop
        [PSCustomObject]@{
            DisplayName        = $mb.DisplayName
            UserPrincipalName  = $mb.UserPrincipalName
            PrimarySmtpAddress = $mb.PrimarySmtpAddress
            TimeZone           = $regional.TimeZone
            Language           = $regional.Language
            DateFormat         = $regional.DateFormat
            TimeFormat         = $regional.TimeFormat
        }
    } catch {
        Write-Warning "Could not read regional configuration for $($mb.UserPrincipalName): $($_.Exception.Message)"
    }
}

$outDir = Split-Path -Path $OutputPath -Parent
if ($outDir -and -not (Test-Path -Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }

$rows | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8 -Force
Write-Output "Exported $($rows.Count) mailbox time zone records to: $OutputPath"
