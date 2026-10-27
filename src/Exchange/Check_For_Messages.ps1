<#
.SYNOPSIS
Check a list of message IDs against a CSV of messages and report which IDs exist in the CSV.

.DESCRIPTION
v2.0.0: this script used to prompt for every input with Read-Host, which
made it unusable non-interactively and untestable. It now takes all inputs
as parameters and fails fast on a missing file, an empty CSV, or a missing
column, instead of looping on "Try again?" prompts.

- The source CSV is expected to have a MessageId column; use
  -MessageIdColumn to name another column.
- The ID file may be comma- or newline-separated; IDs are trimmed and
  de-duplicated before comparison.
- With -Save, found rows are written as CSV and missing IDs as text under
  data/output (or -OutputDir), with a timestamped filename.

.EXAMPLE
pwsh -File src/Exchange/Check_For_Messages.ps1 -CsvPath data/input/messages.csv -IdFilePath data/input/ids.txt -Save
#>
#requires -Version 7.2

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true, HelpMessage='Path to the source CSV file (contains message data)')]
    [string]$CsvPath,

    [Parameter(Mandatory=$true, HelpMessage='Path to the text file containing message IDs (comma or newline separated)')]
    [string]$IdFilePath,

    [Parameter(HelpMessage='Column in the CSV that contains the message id (case-sensitive as in CSV)')]
    [string]$MessageIdColumn,

    [Parameter(HelpMessage='Directory for saved results; defaults to data/output under the repo root')]
    [string]$OutputDir,

    [switch]$Save
)

Write-Host "Check_For_Messages: Compare messageID list against a CSV" -ForegroundColor Cyan

if (-not (Test-Path -Path $CsvPath)) { throw "CSV file not found: $CsvPath" }
$csv = Import-Csv -Path $CsvPath -ErrorAction Stop
if ($csv.Count -eq 0) { throw "CSV seems empty: $CsvPath" }

# Determine MessageId column
$cols = $csv[0].PSObject.Properties | ForEach-Object { $_.Name }
$matchCol = $MessageIdColumn
if (-not $matchCol) {
    $matchCol = $cols | Where-Object { $_ -match '^(?i:messageid)$' } | Select-Object -First 1
}
if (-not $matchCol) { throw "Could not find a 'MessageId' column automatically. Available columns: $($cols -join ', '). Pass -MessageIdColumn." }
if (-not ($cols -contains $matchCol)) { throw "Column '$matchCol' not found in CSV. Available columns: $($cols -join ', ')." }
Write-Host "Using column: $matchCol" -ForegroundColor Green

if (-not (Test-Path -Path $IdFilePath)) { throw "ID file not found: $IdFilePath" }
$raw = Get-Content -Path $IdFilePath -Raw -ErrorAction Stop

# Parse IDs (v2.0.0: split on commas AND newlines. v1 split on commas
# only when a comma was present, so a file with both separators merged
# "id2`n id3" into one bogus id and silently under-counted matches.)
$ids = $raw -split '[,\r\n]+'
$ids = $ids | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique

Write-Host "Parsed $($ids.Count) unique IDs from $IdFilePath" -ForegroundColor Cyan

# Build lookup from CSV
$csvIds = @{ }
foreach ($row in $csv) {
    $val = $row.$matchCol
    if ($null -ne $val) { $csvIds[$val.ToString().Trim()] = $row }
}

# Compare
$found = @()
$missing = @()
foreach ($id in $ids) {
    if ($csvIds.ContainsKey($id)) { $found += $id } else { $missing += $id }
}

Write-Host "Results:" -ForegroundColor Cyan
Write-Host "  Total IDs checked: $($ids.Count)"
Write-Host "  Found in CSV: $($found.Count)" -ForegroundColor Green
Write-Host "  Missing from CSV: $($missing.Count)" -ForegroundColor Yellow

if ($missing.Count -gt 0) {
    Write-Host "Missing sample (first 10):" -ForegroundColor Yellow
    $missing | Select-Object -First 10 | ForEach-Object { Write-Host "  $_" }
}

if ($Save) {
    try {
        if ($OutputDir) {
            $dataOutDir = $OutputDir
        } else {
            $repoRoot = Resolve-Path -Path (Join-Path $PSScriptRoot '../..') | Select-Object -ExpandProperty Path
            $dataOutDir = Join-Path $repoRoot 'data/output'
        }
        if (!(Test-Path -Path $dataOutDir)) { New-Item -ItemType Directory -Path $dataOutDir -Force | Out-Null }
        $base = [System.IO.Path]::GetFileNameWithoutExtension($CsvPath)
        $ts = (Get-Date).ToString('yyyyMMdd_HHmm')
        $foundPath = Join-Path $dataOutDir "$(${base})_found_$ts.csv"
        $missingPath = Join-Path $dataOutDir "$(${base})_missing_$ts.txt"

        $foundRows = $found | ForEach-Object { $csvIds[$_] }
        if ($foundRows.Count -gt 0) { $foundRows | Export-Csv -Path $foundPath -NoTypeInformation -Encoding UTF8 -Force; Write-Host "Saved found rows to: $foundPath" -ForegroundColor Green } else { Write-Host "No found rows to save." }

        if ($missing.Count -gt 0) { $missing -join "`r`n" | Set-Content -Path $missingPath -Encoding UTF8 -Force; Write-Host "Saved missing IDs to: $missingPath" -ForegroundColor Green }
    } catch {
        Write-Warning "Failed to save results: $($_.Exception.Message)"
    }
}

Write-Host "Done." -ForegroundColor Cyan
