<#
.SYNOPSIS
    Debug helper: inspect the attachments on one Graph message.

.DESCRIPTION
    v2.0.0: this was a one-off script at the repo root (debug_attachment.ps1)
    with a hard-coded mailbox. It now lives in examples/ and takes the
    mailbox and message id as parameters, so it cannot run against the
    wrong mailbox by accident. It is a diagnostic only; it changes nothing.

.EXAMPLE
    pwsh -File examples/Debug-Attachment.ps1 -MailboxEmail user@contoso.com -MessageId AAMk...
#>
#requires -Version 7.2

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$MailboxEmail,

    [Parameter(Mandatory=$true)]
    [string]$MessageId
)

. (Join-Path $PSScriptRoot '..' 'src' 'Utils' 'Initialize-Environment.ps1')
$envStatus = Connect-ExchangeEnvironment -GraphScopes @('User.Read.All', 'Mail.Read.Shared') -ErrorAction Stop
if (-not $envStatus.Graph.Connected) { throw "Microsoft Graph is not connected: $($envStatus.Graph.Message)" }

$TargetUser = Get-MgUser -Filter "mail eq '$MailboxEmail'" -ConsistencyLevel eventual -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $TargetUser) { throw "No user found for mailbox: $MailboxEmail" }

Write-Host "Fetching attachments for $MessageId"
$atts = Get-MgUserMessageAttachment -UserId $TargetUser.Id -MessageId $MessageId

foreach ($a in $atts) {
    Write-Host "Name: $($a.Name)"
    Write-Host "Type: $($a.OdataType)"
    Write-Host "ContentType: $($a.ContentType)"

    if ($a.AdditionalProperties.ContainsKey("contentBytes")) {
        Write-Host "Has 'contentBytes' in AdditionalProperties"
        $len = $a.AdditionalProperties["contentBytes"].Length
        Write-Host "Length: $len"
    } elseif ($a.ContentBytes) {
        Write-Host "Has .ContentBytes property"
        Write-Host "Length: $($a.ContentBytes.Length)"
    } else {
        Write-Host "NO CONTENT BYTES FOUND"
        $a | Format-List * | Out-String | Write-Host
    }
}
