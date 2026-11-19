#Requires -Version 7.5
param([Parameter(Mandatory)] [string]$Executable)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$exe = (Resolve-Path -LiteralPath $Executable).Path
$legacy = Join-Path $PSScriptRoot '../../scripts/statistica.ps1'
$work = Join-Path ([IO.Path]::GetTempPath()) ('statistica-csv-test-' + [guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($work)
$spec = Join-Path $work 'sample.json'
$inputSta = Join-Path $work 'sample.sta'
$oldCsv = Join-Path $work 'old.csv'
$newCsv = Join-Path $work 'new.csv'

function Assert-CsvMatch {
    if ((Get-FileHash -LiteralPath $oldCsv).Hash -ne (Get-FileHash -LiteralPath $newCsv).Hash) {
        throw 'The published EXE did not produce the same UTF-8 CSV as the PowerShell CLI.'
    }
}

try {
    @'
{
  "columns": [
    { "name": "Measurement", "type": "double" },
    { "name": "Score", "type": "double" },
    { "name": "Label", "type": "text", "length": 80 },
    { "name": "Count", "type": "double" }
  ],
  "rows": [
    [ 1.25, 0, "alpha,beta", 1 ],
    [ null, 0.12345678912345678, "café \"γ\"", 8 ],
    [ 7.125, -42, null, null ]
  ]
}
'@ | Set-Content -LiteralPath $spec -Encoding utf8
    & $legacy CreateSta -SpecPath $spec -OutputPath $inputSta -Backend Library | Out-Null
    $original = (Get-FileHash -LiteralPath $inputSta).Hash

    $old = & $legacy Inspect -InputPath $inputSta -Backend Library | ConvertFrom-Json -AsHashtable
    $new = & $exe inspect --input-path $inputSta --backend Library | ConvertFrom-Json -AsHashtable
    if ($LASTEXITCODE -ne 0) { throw 'Inspect failed.' }
    foreach ($column in $new.columns) { $column.missing_code = [double]$column.missing_code }
    $oldJson = $old | ConvertTo-Json -Depth 8 -Compress
    $newJson = $new | ConvertTo-Json -Depth 8 -Compress
    if ($oldJson -ne $newJson) {
        throw "The published EXE inspection differs from the PowerShell CLI.`nLegacy: $oldJson`nNative: $newJson"
    }

    & $legacy ExportCsv -InputPath $inputSta -OutputPath $oldCsv -Backend Library | Out-Null
    & $exe export-csv --input-path $inputSta --output-path $newCsv --backend Library | Out-Null
    if ($LASTEXITCODE -ne 0) { throw 'Export failed.' }
    Assert-CsvMatch

    & $exe export-csv --input-path $inputSta --output-path $newCsv --backend Library 2>$null | Out-Null
    if ($LASTEXITCODE -eq 0) { throw 'Export unexpectedly overwrote an existing CSV.' }
    Assert-CsvMatch

    & $legacy ExportCsv -InputPath $inputSta -OutputPath $oldCsv -Backend Application -Columns '3,2,1,2' -FirstRow 2 -LastRow 3 -DisplayText -Force | Out-Null
    & $exe export-csv --input-path $inputSta --output-path $newCsv --backend Application --columns '3,2,1,2' --first-row 2 --last-row 3 --display-text --force | Out-Null
    if ($LASTEXITCODE -ne 0) { throw 'Display text export failed.' }
    Assert-CsvMatch

    & $exe export-csv --input-path $inputSta --output-path $newCsv --backend Library --first-row 9 --force 2>$null | Out-Null
    if ($LASTEXITCODE -eq 0) { throw 'Export unexpectedly accepted an out-of-range row.' }
    Assert-CsvMatch

    if ((Get-FileHash -LiteralPath $inputSta).Hash -ne $original) { throw 'Export changed the source spreadsheet.' }

    $monthlySta = Join-Path $PSScriptRoot '../fixtures/PR_DWLS_025_005_065_20.sta'
    $monthlyHash = (Get-FileHash -LiteralPath $monthlySta).Hash
    foreach ($displayText in $false, $true) {
        $nativeOptions = @()
        if ($displayText) { $nativeOptions += '--display-text' }
        & $legacy ExportCsv -InputPath $monthlySta -OutputPath $oldCsv -Backend Library -Columns '37,35,2,35' -LastRow 3 -DisplayText:$displayText -Force | Out-Null
        & $exe export-csv --input-path $monthlySta --output-path $newCsv --backend Library --columns '37,35,2,35' --last-row 3 --force @nativeOptions | Out-Null
        if ($LASTEXITCODE -ne 0) { throw 'Monthly fixture export failed.' }
        Assert-CsvMatch

        $rows = @(Import-Csv -LiteralPath $newCsv)
        $expectedDate = if ($displayText) { 'Feb-2008' } else { '103' }
        if ($rows.Count -ne 3 -or ($rows[0].PSObject.Properties.Name -join ',') -cne 'case_index,v37,v35,v2' -or
            $rows[1].v37 -cne $expectedDate -or $rows[2].v35 -cne '') {
            throw 'Monthly fixture export changed column order, date labels or missing values.'
        }
    }
    if ((Get-FileHash -LiteralPath $monthlySta).Hash -ne $monthlyHash) { throw 'Export changed the monthly fixture.' }

    if (Get-ChildItem -LiteralPath $work -Filter '.statistica-*' -Force) { throw 'A staging file survived cleanup.' }
    'Inspect, CSV byte parity, overwrite, rollback and source integrity passed.'
} finally {
    foreach ($file in [IO.Directory]::EnumerateFiles($work)) { [IO.File]::Delete($file) }
    [IO.Directory]::Delete($work)
}
exit 0
