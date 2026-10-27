<#
.SYNOPSIS
Interactively export a mailbox to markdown files and a CSV report.

.NOTES
Requires Microsoft Graph PowerShell SDK (Connect-MgGraph, Get-MgUser, Get-MgUserMailFolderMessage, etc.).
#>
#requires -Version 7.2

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true, HelpMessage = "Enter the mailbox email to organize (e.g. user@contoso.com)")]
    [string]$MailboxEmail,

    [Parameter(Mandatory = $false, HelpMessage = "Name of the folder to start exporting from. Defaults to 'Inbox'.")]
    [string]$StartFolder = 'Inbox',

    [Parameter(Mandatory = $false, HelpMessage = "Maximum number of messages to retrieve per folder. Defaults to 50.")]
    [int]$Limit = 50,

    [Parameter(Mandatory = $false, HelpMessage = "Recursively export sub-folders.")]
    [switch]$Recurse,

    [Parameter(Mandatory = $false, HelpMessage = "Include attachments in the export.")]
    [switch]$IncludeAttachments,

    [Parameter(Mandatory = $false, HelpMessage = "Download attachment files to disk.")]
    [switch]$DownloadAttachments
)

# --- Helper Functions ---

# Helper: Convert simple HTML to plain text
function Convert-HtmlToText {
	param ([string]$HtmlContent)
	if ([string]::IsNullOrWhiteSpace($HtmlContent)) { return "" }

	$text = $HtmlContent -replace '<br[^>]*>', "`r`n"
	$text = $text -replace '<[^>]+>', ""
	# Replace common HTML entities
	# cspell:ignore nbsp
	$text = $text -replace '&nbsp;', ' '  # non-breaking space
	$text = $text -replace '&amp;', '&'
	$text = $text -replace '&lt;', '<'
	$text = $text -replace '&gt;', '>'

	return $text -replace "(?m)^\s*\r?\n", ""
}

# Helper: Make filename safe
function Get-SafeFilename {
	param ([string]$Subject)
	if ($null -eq $Subject) { $Subject = '' }
	$safe = $Subject -replace '[\\/:*?"<>|]', '_'
	$safe = $safe -replace "[\r\n\t]", ' '
	if ($safe.Length -gt 50) { $safe = $safe.Substring(0,50) }
	return $safe.Trim()
}

# Return a filename safe for attachments, truncated and made unique in $Dir
function Get-SafeAttachmentFilename {
    param(
        [Parameter(Mandatory=$true)][string]$Dir,
        [Parameter(Mandatory=$true)][string]$OriginalName,
        [int]$MaxBaseLength = 100
    )

    if ($null -eq $OriginalName) { $OriginalName = 'attachment' }
    # sanitize
    $safe = $OriginalName -replace '[\\/:*?"<>|]', '_'
    $safe = $safe -replace "[\r\n\t]", ' '

    $ext = [System.IO.Path]::GetExtension($safe)
    $nameNoExt = [System.IO.Path]::GetFileNameWithoutExtension($safe)
    if ($nameNoExt.Length -gt $MaxBaseLength) { $nameNoExt = $nameNoExt.Substring(0,$MaxBaseLength) }

    $candidate = if ($ext) { "$nameNoExt$ext" } else { $nameNoExt }
    $i = 1
    while (Test-Path -LiteralPath (Join-Path $Dir $candidate)) {
        $candidate = if ($ext) { "$nameNoExt`_$i$ext" } else { "$nameNoExt`_$i" }
        $i++
        if ($i -gt 9999) { break }
    }
    return $candidate
}

# --- Load Utilities ---
try {
    . "$PSScriptRoot/../Utils/Initialize-Environment.ps1"
    . "$PSScriptRoot/../Utils/LoggingHelpers.ps1"
} catch {
    Write-Error "Failed to load Utils. Ensure '..\Utils\Initialize-Environment.ps1' and 'LoggingHelpers.ps1' exist."
    return
}

# --- Initialize Environment ---
# Connect to Graph (User.Read.All and Mail.Read.Shared required)
$EnvStatus = Connect-ExchangeEnvironment -GraphScopes "User.Read.All","Mail.Read.Shared"
if (-not $EnvStatus.Graph.Connected) {
    Write-Error "Failed to connect to Microsoft Graph: $($EnvStatus.Graph.Message)"
    return
}

# --- Setup Logging ---
$repoRoot = Resolve-Path -Path (Join-Path $PSScriptRoot '../..') | Select-Object -ExpandProperty Path
$logDir = Join-Path $repoRoot 'logs/Export-MailboxMessages'
if (!(Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }

$mailboxSafe = Get-SafeFilename -Subject $MailboxEmail
$logName = "$($mailboxSafe)_export_messages_$(Get-Date -Format 'yyyyMMdd_HHmm').log"

Set-LogFile -Path (Join-Path $logDir $logName)
Write-LogMessage -Level INFO -Message "Starting Export-MailboxMessages for $MailboxEmail"

Write-LogMessage -Level INFO -Message "Looking up user by email address: $MailboxEmail"

$filter1 = "mail eq '$MailboxEmail'"
$TargetUser = Get-MgUser -Filter $filter1 -ConsistencyLevel eventual -CountVariable c -ErrorAction SilentlyContinue | Select-Object -First 1

if (-not $TargetUser) {
	Write-LogMessage -Level WARN -Message "Primary mail lookup failed. Checking proxy addresses..."
	$filter2 = "proxyAddresses/any(p:p eq 'smtp:$MailboxEmail')"
	$TargetUser = Get-MgUser -Filter $filter2 -ConsistencyLevel eventual -CountVariable c -ErrorAction SilentlyContinue | Select-Object -First 1
}

if ($TargetUser) {
	Write-LogMessage -Level INFO -Message "Found Mailbox: $($TargetUser.DisplayName) (ObjectId: $($TargetUser.Id))"
} else {
	Write-Error "CRITICAL: Still could not find a user with email '$MailboxEmail'."
    Write-LogMessage -Level ERROR -Message "CRITICAL: Still could not find a user with email '$MailboxEmail'."
	return
}

# Prepare CSV collection
$Global:ReportList = [System.Collections.Generic.List[PSObject]]::new()

function Export-MessagesRecursive {
	param (
		[string]$FolderId,
		[string]$FolderPath
	)

	Write-LogMessage -Level INFO -Message "Processing Folder: $FolderPath"

    # Add 'HasAttachments' to property selection
	$Props = @('Id','ReceivedDateTime','Subject','From','ToRecipients','Body','HasAttachments')
    
    if ($Limit -gt 0) {
	    $Msgs = Get-MgUserMailFolderMessage -UserId $TargetUser.Id -MailFolderId $FolderId -Property $Props -PageSize $Limit -Top $Limit -All:$false -ErrorAction SilentlyContinue
    } else {
        # Limit 0 or less implies All
        $Msgs = Get-MgUserMailFolderMessage -UserId $TargetUser.Id -MailFolderId $FolderId -Property $Props -PageSize 100 -All -ErrorAction SilentlyContinue
    }

	foreach ($msg in $Msgs) {
		# cspell:ignore rcvd
		try { $rcvd = $msg.ReceivedDateTime } catch { $rcvd = $null }

		if ($msg.Body -and $msg.Body.ContentType -and ($msg.Body.ContentType -eq 'html')) {
			$BodyText = Convert-HtmlToText -HtmlContent $msg.Body.Content
		} else { $BodyText = $msg.Body.Content }

		if ($msg.From) { $FromAddress = $msg.From.EmailAddress.Address } else { $FromAddress = 'Unknown' }
		$Recipients = @()
		if ($msg.ToRecipients) { $Recipients = $msg.ToRecipients | ForEach-Object { $_.EmailAddress.Address } }
		$Recipients = ($Recipients -join '; ')

        $AttachmentNames = ""
        if (($IncludeAttachments -or $DownloadAttachments) -and $msg.HasAttachments) {
            try {
                        $attachments = Get-MgUserMessageAttachment -UserId $TargetUser.Id -MessageId $msg.Id -ErrorAction Stop
                        if ($attachments) {
                            $names = $attachments | ForEach-Object { $_.Name }
                            $AttachmentNames = ($names -join '; ')

                    if ($DownloadAttachments) {
                         # Structure: attachments/<Sender>/<YYYYMMDD_HHmm>_<Subject>/
                         $safeSender = Get-SafeFilename -Subject $FromAddress
                         
                         $dateStr = "NoDate"
                         if ($rcvd -is [DateTime]) { $dateStr = $rcvd.ToString("yyyyMMdd_HHmm") }
                         elseif ($rcvd) { $dateStr = Get-SafeFilename -Subject "$rcvd" }

                         # Use only the date/time for the message folder to keep paths shorter
                         $folderName = $dateStr

                         $senderDir = Join-Path $AttachmentsBaseDir $safeSender
                         if (!(Test-Path $senderDir)) { New-Item -ItemType Directory -Path $senderDir -Force | Out-Null }

                         # Ensure uniqueness if a folder with the same timestamp already exists
                         $origFolderName = $folderName
                         $idx = 1
                         $msgAttDir = Join-Path $senderDir $folderName
                         while (Test-Path $msgAttDir) {
                             $folderName = "${origFolderName}_$idx"
                             $msgAttDir = Join-Path $senderDir $folderName
                             $idx++
                             if ($idx -gt 9999) { break }
                         }
                         New-Item -ItemType Directory -Path $msgAttDir -Force | Out-Null
                         
                         foreach ($att in $attachments) {
                             # Try to get content bytes
                             $finalBytes = $null
                             
                             if ($att.ContentBytes) {
                                 $finalBytes = $att.ContentBytes
                             } elseif ($att.AdditionalProperties -and $att.AdditionalProperties.ContainsKey('contentBytes')) {
                                 $finalBytes = $att.AdditionalProperties['contentBytes']
                             }
                             
                             if ($finalBytes) {
                                  # In SDK, this is likely byte[]. If string (base64), convert.
                                  if ($finalBytes -is [string]) {
                                      $bytes = [System.Convert]::FromBase64String($finalBytes)
                                  } else {
                                      $bytes = $finalBytes
                                  }
                                  
                                  # Create a filename that respects path length limits and is unique
                                  $maxFullPath = 240
                                  $baseLen = 100
                                  $fName = Get-SafeAttachmentFilename -Dir $msgAttDir -OriginalName $att.Name -MaxBaseLength $baseLen
                                  $fPath = Join-Path $msgAttDir $fName
                                  # If full path too long, reduce base length until it fits
                                  while ($fPath.Length -gt $maxFullPath -and $baseLen -gt 10) {
                                      $baseLen = [int]([Math]::Floor($baseLen * 0.8))
                                      $fName = Get-SafeAttachmentFilename -Dir $msgAttDir -OriginalName $att.Name -MaxBaseLength $baseLen
                                      $fPath = Join-Path $msgAttDir $fName
                                  }
                                  try {
                                      [System.IO.File]::WriteAllBytes($fPath, $bytes)
                                  } catch {
                                      Write-LogMessage -Level WARN -Message "Failed to write attachment $fName : $($_.Exception.Message)"
                                  }
                             }
                         }
                    }
                }
            } catch {
                Write-LogMessage -Level WARN -Message "Failed to fetch attachments for MsgId $($msg.Id): $($_.Exception.Message)"
            }
        }

		$csvRow = [PSCustomObject]@{
			MessageId        = $msg.Id
			ReceivedDateTime = $rcvd
			FolderPath       = $FolderPath
			From             = $FromAddress
			To               = $Recipients
			Subject          = $msg.Subject
            HasAttachments   = $msg.HasAttachments
            AttachmentNames  = $AttachmentNames
			BodyCleaned      = $BodyText
		}
		$Global:ReportList.Add($csvRow)
	}

	if ($Recurse) {
		try {
			$SubFolders = Get-MgUserMailFolderChildFolder -UserId $TargetUser.Id -MailFolderId $FolderId -PageSize 100 -All -ErrorAction SilentlyContinue
		} catch { $SubFolders = @() }

		foreach ($sub in $SubFolders) {
			Export-MessagesRecursive -FolderId $sub.Id -FolderPath "$FolderPath\$($sub.DisplayName)"
		}
	}
}

# Prepare Output Directories
$repoRoot = Resolve-Path -Path (Join-Path $PSScriptRoot '../..') | Select-Object -ExpandProperty Path
$mailboxSafe = Get-SafeFilename -Subject $MailboxEmail
$DataOutDir = Join-Path $repoRoot "data/output/$mailboxSafe"
if (!(Test-Path -Path $DataOutDir)) { New-Item -ItemType Directory -Path $DataOutDir -Force | Out-Null }

$AttachmentsBaseDir = Join-Path $DataOutDir "attachments"
if ($DownloadAttachments -and !(Test-Path $AttachmentsBaseDir)) { New-Item -ItemType Directory -Path $AttachmentsBaseDir -Force | Out-Null }

# Start from specified folder
$RootFolder = Get-MgUserMailFolder -UserId $TargetUser.Id -Filter "displayName eq '$StartFolder'" -Top 1 -ErrorAction SilentlyContinue | Select-Object -First 1

if ($RootFolder) {
	Export-MessagesRecursive -FolderId $RootFolder.Id -FolderPath $StartFolder

    # Write CSV into repository data\output\<SafeMailboxName> using mailbox-based filename
	try {
		$outName = "$($mailboxSafe)_email_log_$((Get-Date).ToString('yyyyMMdd_HHmm')).csv"
		$outPath = Join-Path -Path $DataOutDir -ChildPath $outName
		
		Write-LogMessage -Level INFO -Message "Writing CSV file to: $outPath"
		$Global:ReportList | Export-Csv -Path $outPath -NoTypeInformation -Encoding UTF8 -Force
		Write-LogMessage -Level INFO -Message "Done! CSV located at: $outPath"
	} catch {
		Write-Error "Failed to write CSV to data/output: $($_.Exception.Message)"
		Write-LogMessage -Level ERROR -Message "Failed to write CSV: $($_.Exception.Message)"
	}
} else {
	Write-Error "Inbox not found for user $MailboxEmail"
	Write-LogMessage -Level ERROR -Message "Inbox not found for user $MailboxEmail"
}
