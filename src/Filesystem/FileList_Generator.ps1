<#
.SYNOPSIS
    Generate a file list (placeholder).
.DESCRIPTION
    Walks a directory and outputs a CSV/list of files.
#>
#requires -Version 7.2

param(
    [Parameter(Mandatory=$false)]
    [string]$Path = ".",
    [Parameter(Mandatory=$false)]
    [string]$OutputFile = "data/output/file_list.csv"
)

Write-Output "[INFO] FileList_Generator - scanning: $Path"

if (-not (Test-Path -LiteralPath $Path)) {
    Write-Error "Source path does not exist: $Path"
    exit 1
}

$outDir = Split-Path -Path $OutputFile -Parent
if ($outDir -and -not (Test-Path $outDir)) {
    New-Item -ItemType Directory -Path $outDir -Force | Out-Null
}

try {
    $files = Get-ChildItem -LiteralPath $Path -Recurse -File -ErrorAction Stop
    $count = $files.Count
    Write-Output "[INFO] Found $count files. Exporting..."

    $files | Select-Object FullName, Length, LastWriteTime, CreationTime, Extension | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8 -Force
    Write-Output "[INFO] Done. Output saved to: $OutputFile"
}
catch {
    Write-Error "An error occurred: $($_.Exception.Message)"
}
