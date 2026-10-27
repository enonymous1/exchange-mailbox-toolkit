<#
.SYNOPSIS
    Export a shared mailbox's calendar to CSV using Microsoft Graph.
.DESCRIPTION
    Prompts for a shared mailbox email address, ensures the environment (Graph) is
    connected via the `Connect-ExchangeEnvironment` helper, then exports the mailbox's calendar
    items for a date range (6 months back to 1 month forward) to `data/output`.
.EXAMPLE
    # Interactive
    pwsh -File src\Exchange\Export-SharedCalendar.ps1

    # Non-interactive (pass email and optional output path)
    pwsh -File src\Exchange\Export-SharedCalendar.ps1 -SharedMailboxEmail "shared@contoso.com"
#>
#requires -Version 7.2

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [string]$SharedMailboxEmail,

    [Parameter(Mandatory=$false)]
    [string]$OutputPath,

    [Parameter(Mandatory=$false)]
    [string]$OutDir,

    [Parameter(Mandatory=$false)]
    [ValidateSet('Shared','User','All')]
    [string]$MailboxType = 'Shared',

    [Parameter(Mandatory=$false)]
    [switch]$InternalOnly,

    [Parameter(Mandatory=$false)]
    [switch]$ExternalOnly,

    [Parameter(Mandatory=$false)]
    [string[]]$Domains,

    [Parameter(Mandatory=$false)]
    [switch]$ListMailboxes,

    [Parameter(Mandatory=$false)]
    [switch]$WhatIf
)

# --- CONFIGURATION ---
if (-not $SharedMailboxEmail) {
    if ($ListMailboxes) {
        # We'll browse for shared mailboxes after confirming Graph connection
        $BrowseSharedMailboxes = $true
    }
    else {
        $SharedMailboxEmail = Read-Host -Prompt "Enter the Shared Mailbox Email Address (or press Enter to browse)"
        if ([string]::IsNullOrWhiteSpace($SharedMailboxEmail)) { $BrowseSharedMailboxes = $true }
    }
}

if (-not $BrowseSharedMailboxes -and [string]::IsNullOrWhiteSpace($SharedMailboxEmail)) {
    Write-Error "Email address cannot be empty."
    exit 1
}

# Date range: 6 months back, 1 month forward
$StartDate = (Get-Date).AddMonths(-6).ToString('yyyy-MM-ddTHH:mm:ssZ')
$EndDate   = (Get-Date).AddMonths(1).ToString('yyyy-MM-ddTHH:mm:ssZ')

# Repo root and output directory (prefer -OutDir, else repo data/output)
$RepoRoot = Resolve-Path -Path (Join-Path $PSScriptRoot '../..')
if ($OutDir) {
    $OutputDirCandidate = Resolve-Path -Path $OutDir -ErrorAction SilentlyContinue
    if (-not $OutputDirCandidate) { New-Item -Path $OutDir -ItemType Directory -Force | Out-Null; $OutputDirCandidate = Resolve-Path -Path $OutDir }
    $OutputDir = $OutputDirCandidate.Path
}
else {
    $OutputDir = Join-Path $RepoRoot 'data/output'
    if (-not (Test-Path $OutputDir)) { New-Item -Path $OutputDir -ItemType Directory -Force | Out-Null }
}

# If caller passed an explicit OutputPath, leave it (resolve relative paths); otherwise each user export will get its own filename in $OutputDir
if ($OutputPath) {
    $Resolved = Resolve-Path -Path $OutputPath -ErrorAction SilentlyContinue
    if ($Resolved) { $OutputPath = $Resolved.Path }
    else { $OutputPath = Join-Path $OutputDir $OutputPath }
}

# Initialize result containers
$ExportSuccess = @()
$ExportFailures = @()

# Setup logging (use repo logs folder)
$LogDir = Join-Path $RepoRoot 'logs'
if (-not (Test-Path $LogDir)) { New-Item -Path $LogDir -ItemType Directory -Force | Out-Null }
$LogFile = Join-Path $LogDir ("export_shared_calendar_{0:yyyyMMdd}.log" -f (Get-Date))

# Load the logging helpers (v2.0.0: the old Helpers.ps1 file never existed).
$loggingHelpersPath = Join-Path $PSScriptRoot '../Utils/LoggingHelpers.ps1'
if (Test-Path $loggingHelpersPath) {
    try { . $loggingHelpersPath } catch { Write-Warning "Could not dot-source LoggingHelpers.ps1: $($_.Exception.Message)" }
}

# Configure helper log file if available
if (Get-Command -Name Set-LogFile -ErrorAction SilentlyContinue) { Set-LogFile -Path $LogFile }

# Use Write-LogMessage (helper) instead of a local Append-LogFile function; Set-LogFile already configured above to point to $LogFile

# Log start
if (Get-Command -Name Write-LogMessage -ErrorAction SilentlyContinue) { Write-LogMessage -Message "Starting Export-SharedCalendar for $SharedMailboxEmail" -Level 'INFO' } else { Write-Verbose "Starting Export-SharedCalendar for $SharedMailboxEmail" }
Write-LogMessage -Message "Starting Export-SharedCalendar for $SharedMailboxEmail" -Level 'INFO'

# --- USE SHARED ENV HELPER FOR AUTHENTICATION ---
$RequiredScopes = @('User.Read.All','Calendars.Read','Calendars.Read.Shared')

# Call the helper function we added in Utils
try {
    # Dot-source the helper (if not imported as a module)
    if (-not (Get-Command -Name Connect-ExchangeEnvironment -ErrorAction SilentlyContinue)) {
        . "$PSScriptRoot/../Utils/Connect-ExchangeEnvironment.ps1"
    }
}
catch {
    Write-Error "Failed to load Connect-ExchangeEnvironment helper: $($_.Exception.Message)"
    exit 1
}

# Establish connections (Connect-ExchangeEnvironment returns a status object)
$status = Connect-ExchangeEnvironment -GraphScopes $RequiredScopes -LogFile $LogFile -Verbose
if ($status.Graph.Connected) {
    if (Get-Command -Name Write-LogMessage -ErrorAction SilentlyContinue) { Write-LogMessage -Message "Microsoft Graph connection established for $($status.Graph.Account)" -Level 'INFO' } else { Write-Verbose "Microsoft Graph connected: $($status.Graph.Account)" }
    Write-LogMessage -Message "Microsoft Graph connected: $($status.Graph.Account)" -Level 'INFO'
}
else {
    if (Get-Command -Name Write-LogMessage -ErrorAction SilentlyContinue) { Write-LogMessage -Message "Failed to establish Microsoft Graph connection" -Level 'ERROR' } else { Write-Verbose "Failed to establish Microsoft Graph connection" }
    if ($status.Errors) { $status.Errors | ForEach-Object { if (Get-Command -Name Write-LogMessage -ErrorAction SilentlyContinue) { Write-LogMessage -Message $_ -Level 'ERROR' } } }
    Write-LogMessage -Message "Aborting due to Graph connection failure" -Level 'ERROR'
    Write-Error "Could not establish Microsoft Graph connection. Aborting."
    exit 1
}

# Helper: parse selection string like '1,3-5' into 1-based indices
function Get-SelectionIndices {
    param(
        [string]$Selection,
        [int]$Max
    )
    if (-not $Selection) { return @() }

    # Normalize and diagnostic
    $orig = $Selection
    $in = $Selection.Trim()
    $in = $in -replace '\p{Z}',''
    Write-Verbose "Get-SelectionIndices: raw='$orig' trimmed='$in' Max=$Max"

    if ($in -match '^(?i)all$') { return 1..$Max }

    # Simple single-number fast path
    if ($in -match '^[0-9]+$') { return ,([int]$in) }

    $parts = $in -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }
    $result = [System.Collections.Generic.List[int]]::new()
    foreach ($p in $parts) {
        if ($p -match '^([0-9]+)-([0-9]+)$') {
            $start = [int]$Matches[1]; $end = [int]$Matches[2]
            if ($start -le $end) {
                for ($i = $start; $i -le $end; $i++) { $result.Add($i) }
            }
        }
        elseif ($p -match '^[0-9]+$') { $result.Add([int]$p) }
        else { Write-Verbose "Get-SelectionIndices: ignoring non-numeric part: '$p'" }
    }
    $out = $result | Where-Object { $_ -ge 1 -and $_ -le $Max } | Sort-Object -Unique
    return ,$out
}

function Export-CalendarForUser {
    param([pscustomobject]$User)
    $email = $User.Mail
    $id = $User.Id
    $clean = ($email -replace '[^a-zA-Z0-9]', '_')

    # Determine output file: if caller specified a single explicit OutputPath and only one target, use it; otherwise generate per-user file in $OutputDir
    if ($OutputPath -and $SelectedUsers.Count -eq 1) {
        $out = $OutputPath
    }
    else {
        $out = Join-Path $OutputDir ("Calendar_{0}_{1:yyyyMMddHHmmss}.csv" -f $clean, (Get-Date))
    }

    Write-Verbose "Fetching calendar for $email (ID: $id)..."
    $CalendarParams = @{
        UserId        = $id
        StartDateTime = $StartDate
        EndDateTime   = $EndDate
        All           = $true
        Property      = @('subject','start','end','organizer','isOnlineMeeting','onlineMeetingProvider')
    }

    try {
        $CalendarItems = Get-MgUserCalendarView @CalendarParams -ErrorAction Stop
    }
    catch {
        $reason = $_.Exception.Message
        Write-Warning "Export failed for $($email): $reason"
        if (Get-Command -Name Write-LogMessage -ErrorAction SilentlyContinue) { Write-LogMessage -Message "Export failed for $($email): $reason" -Level 'ERROR' } else { Write-Verbose "Export failed for $($email): $reason" }
        $ExportFailures += [pscustomobject]@{ Mail = $email; Id = $id; Reason = $reason; Time = (Get-Date).ToString('o') }
        return
    }

    if ($CalendarItems) {
        $ExportData = $CalendarItems | ForEach-Object {
            [PSCustomObject]@{
                Subject        = $_.Subject
                Organizer      = $_.Organizer.EmailAddress.Name
                OrganizerEmail = $_.Organizer.EmailAddress.Address
                StartTime      = $_.Start.DateTime
                EndTime        = $_.End.DateTime
                IsOnlineMeeting = $_.IsOnlineMeeting
                Platform       = ($_.OnlineMeetingProvider -join ', ')
            }
        }

        if ($WhatIf) {
            Write-Output "WHATIF: Would export {0} items for {1} to {2}" -f ($ExportData.Count), $email, $out
            if (Get-Command -Name Write-LogMessage -ErrorAction SilentlyContinue) { Write-LogMessage -Message "WHATIF: Would export $($ExportData.Count) items for $email to $out" -Level 'INFO' }
            $ExportSuccess += [pscustomobject]@{ Mail = $email; File = $out; Count = $ExportData.Count; WhatIf = $true; Time = (Get-Date).ToString('o') }
        }
        else {
            try {
                $ExportData | Export-Csv -Path $out -NoTypeInformation -Force
                Write-Output "SUCCESS! Export saved to: $out"
                if (Get-Command -Name Write-LogMessage -ErrorAction SilentlyContinue) { Write-LogMessage -Message "Export saved to: $out" -Level 'INFO' }
                $ExportSuccess += [pscustomobject]@{ Mail = $email; File = $out; Count = $ExportData.Count; WhatIf = $false; Time = (Get-Date).ToString('o') }
            }
            catch {
                $reason = $_.Exception.Message
                Write-Warning "Failed to export CSV for $($email): $reason"
                if (Get-Command -Name Write-LogMessage -ErrorAction SilentlyContinue) { Write-LogMessage -Message "Failed to export CSV for $($email): $reason" -Level 'ERROR' }
                $ExportFailures += [pscustomobject]@{ Mail = $email; Id = $id; Reason = "ExportCSV: $reason"; Time = (Get-Date).ToString('o') }
            }
        }
    }
    else {
        Write-Warning "Access successful, but no calendar items found for $email between $StartDate and $EndDate."
        if (Get-Command -Name Write-LogMessage -ErrorAction SilentlyContinue) { Write-LogMessage -Message "No calendar items found between $StartDate and $EndDate for $email" -Level 'WARN' }
        $ExportSuccess += [pscustomobject]@{ Mail = $email; File = $out; Count = 0; WhatIf = $WhatIf.IsPresent; Time = (Get-Date).ToString('o') }
    }
}

# --- MAIN EXECUTION ---
try {
    # If user wanted to browse shared mailboxes, query Graph for shared user types
    $SelectedUsers = @()
    if ($BrowseSharedMailboxes) {
        Write-Verbose "Retrieving shared mailboxes list..."
        try {
            # Use the new helper to fetch mailbox list (supports Shared/User/All and internal/external filters)
            . "$PSScriptRoot/Get-GraphMailboxes.ps1"
            $shared = Get-GraphMailboxes -Type $MailboxType -InternalOnly:$InternalOnly -ExternalOnly:$ExternalOnly -Domains $Domains -Max 0 | Sort-Object DisplayName
            Write-Verbose "Retrieved $($shared.Count) mailboxes (MailboxType=$MailboxType, InternalOnly=$InternalOnly, ExternalOnly=$ExternalOnly)"
            if (Get-Command -Name Write-LogMessage -ErrorAction SilentlyContinue) { Write-LogMessage -Message "Retrieved $($shared.Count) mailboxes (MailboxType=$MailboxType)" -Level 'INFO' }
        }
        catch {
            Write-Error "Failed to retrieve mailbox list: $($_.Exception.Message)"
            exit 1
        }
        if (-not $shared -or $shared.Count -eq 0) {
            Write-Warning "No shared mailboxes found to select."
            exit 1
        }
        for ($i = 0; $i -lt $shared.Count; $i++) { Write-Output ("[{0}] {1} - {2}" -f ($i+1), $shared[$i].DisplayName, $shared[$i].Mail) }
        $attempts = 0
        while ($attempts -lt 3 -and $SelectedUsers.Count -eq 0) {
            $sel = Read-Host "Select mailbox indices (e.g., 1,3-5), 'all', or enter an email"
            $sel = $sel.Trim()
            if (-not $sel) { Write-Warning "Empty selection; try again."; $attempts++; continue }

            # Normalize input and detect numeric-only selections robustly (remove separators and control chars)
            $selNoSpace = $sel -replace '[\p{Z}\p{C}]',''
            $selDigitsOnly = $selNoSpace -replace '[,\-]',''
            Write-Verbose "Selection input raw='$sel' normalized='$selNoSpace' digitsOnly='$selDigitsOnly'"

            if ($sel -match '^(?i)all$' -or ($selDigitsOnly -match '^[0-9]+$')) {
                # Inline robust numeric parser (avoid edge cases with selection parsing in Unicode/paste scenarios)
                $norm = $sel -replace '[\p{Z}\p{C}]',''
                $parts = $norm -split ',' | Where-Object { $_ -ne '' }
                $parsed = [System.Collections.Generic.List[int]]::new()
                foreach ($p in $parts) {
                    if ($p -match '^([0-9]+)-([0-9]+)$') {
                        $s = [int]$Matches[1]; $e = [int]$Matches[2]
                        if ($s -le $e) { for ($i = $s; $i -le $e; $i++) { $parsed.Add($i) } }
                    }
                    elseif ($p -match '^[0-9]+$') { $parsed.Add([int]$p) }
                    else { Write-Verbose "Numeric parser ignoring part: '$p'" }
                }
                $idx = $parsed | Sort-Object -Unique
                Write-Verbose "Numeric parser yielded: $($idx -join ',') (Count: $($idx.Count)) | selRaw='$sel' normalized='$selNoSpace' digitsOnly='$selDigitsOnly'"

                if (-not $idx -or $idx.Count -eq 0) {
                    Write-Warning "Parsed no indices for numeric selection '$sel'. Please try again."
                    $attempts++; continue
                }

                $outOfRange = $idx | Where-Object { $_ -lt 1 -or $_ -gt $shared.Count }
                if ($outOfRange -and $outOfRange.Count -gt 0) {
                    Write-Warning "Numeric selection contains out-of-range indices: $($outOfRange -join ',') (valid 1..$($shared.Count)). Please try again."
                    $attempts++; continue
                }

                Write-Verbose "Parsed numeric selection indices: $($idx -join ',')"
                foreach ($n in $idx) { $SelectedUsers += $shared[$n - 1] }
                break
            }

            # Non-numeric input: try exact email match (if contains @) or partial name/email match
            if ($sel -match '@') {
                $match = $shared | Where-Object { $_.Mail -ieq $sel }
            }
            else {
                $match = $shared | Where-Object { $_.DisplayName -like "*$sel*" -or $_.Mail -like "*$sel*" }
            }
            if ($match -and $match.Count -gt 0) {
                $SelectedUsers += $match
                break
            }
            Write-Warning "No valid selection found for '$sel'. Please try again."
            $attempts++
        }
        if ($SelectedUsers.Count -eq 0) { Write-Warning "No selection made after multiple attempts; aborting."; exit 1 }
        # Confirm selection with user
        Write-Output "You selected $($SelectedUsers.Count) mailbox(es):"
        $SelectedUsers | ForEach-Object { Write-Output (" - {0} ({1})" -f $_.DisplayName, $_.Mail) }
        $confirm = Read-Host "Proceed to export these mailboxes? (Y/n)"
        if ($confirm -and $confirm.Trim().ToLower() -eq 'n') { Write-Output 'Aborting by user request.'; exit 0 }
    }
    else {
        # fetch single target user by email
        Write-Verbose "Looking up mailbox: $SharedMailboxEmail ..."
        $UserParams = @{ Filter = "mail eq '$SharedMailboxEmail'" }
        $TargetUser = Get-MgUser @UserParams -ErrorAction Stop
        if (-not $TargetUser) { throw "User/Mailbox with email '$SharedMailboxEmail' was not found in Azure AD." }
        $SelectedUsers = @($TargetUser)
    }

    # Loop and export for each selected user
    foreach ($u in $SelectedUsers) {
        Export-CalendarForUser -User $u
    }

    # --- End-of-run summary ---
    $summary = [ordered]@{
        Timestamp      = (Get-Date).ToString('o')
        TotalSelected  = $SelectedUsers.Count
        SuccessCount   = $ExportSuccess.Count
        FailureCount   = $ExportFailures.Count
        Successes      = $ExportSuccess
        Failures       = $ExportFailures
    }
    $summaryFile = Join-Path $LogDir ("export_summary_{0:yyyyMMddHHmmss}.json" -f (Get-Date))
    $summary | ConvertTo-Json -Depth 5 | Set-Content -Path $summaryFile -Force
    if (Get-Command -Name Write-LogMessage -ErrorAction SilentlyContinue) { Write-LogMessage -Message "Export completed: $($summary.SuccessCount) succeeded, $($summary.FailureCount) failed. Summary: $summaryFile" -Level 'INFO' } else { Write-Output "Export completed: $($summary.SuccessCount) succeeded, $($summary.FailureCount) failed. Summary file: $summaryFile" }
    Write-Output "Summary file written to: $summaryFile"
}
catch {
    $errMsg = $_.Exception.Message
    if ($errMsg -match "Access is denied|403") {
        Write-Error "ACCESS DENIED (403)"
        Write-Warning "Graph API cannot open this calendar. This is usually because:"
        Write-Information "1. You do not have 'Full Access' delegation on the shared mailbox."
        Write-Information "2. You haven't waited 30-60 mins after granting permission."
        Write-Information "ACTION: Go to Exchange Admin Center > Recipients > Mailboxes > Delegate > Read and Manage."
        Write-LogMessage -Message "ACCESS DENIED when accessing calendar: $errMsg" -Level 'ERROR'
    }
    else {
        Write-Error "An unexpected error occurred: $errMsg"
        Write-LogMessage -Message "Unexpected error: $errMsg" -Level 'ERROR'
    }
    exit 1
}
