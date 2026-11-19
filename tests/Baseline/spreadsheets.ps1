#Requires -Version 7.5
param([Parameter(Mandatory)] [string]$Executable)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$PSNativeCommandUseErrorActionPreference = $true
$exe = (Resolve-Path -LiteralPath $Executable).Path
$work = Join-Path ([IO.Path]::GetTempPath()) ('statistica-csv-test-' + [guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($work)
$spec = Join-Path $work 'sample.json'
$inputSta = Join-Path $work 'sample.sta'
$outputCsv = Join-Path $work 'sample.csv'

try {
    @"
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
"@ | Set-Content -LiteralPath $spec -Encoding utf8
    & $exe create-sta --spec-path $spec --output-path $inputSta --backend Library | Out-Null
    $original = (Get-FileHash -LiteralPath $inputSta).Hash

    $info = & $exe inspect --input-path $inputSta --backend Library | ConvertFrom-Json
    if ($info.path -cne $inputSta -or $info.cases -ne 3 -or $info.variables -ne 4 -or
        ($info.columns.name -join ',') -cne 'Measurement,Score,Label,Count' -or ($info.columns.type -join ',') -ne '0,0,1,0') {
        throw 'Inspection changed the spreadsheet dimensions or columns.'
    }

    & $exe export-csv --input-path $inputSta --output-path $outputCsv --backend Library | Out-Null
    $expected = @(
        '"case_index","v1","v2","v3","v4"'
        '"1","1.25","0","alpha,beta","1"'
        '"2","","0.12345678912345678","café ""γ""","8"'
        '"3","7.125","-42","",""'
    ) -join [Environment]::NewLine
    $expected += [Environment]::NewLine
    $raw = [IO.File]::ReadAllBytes($outputCsv)
    if ([Convert]::ToHexString($raw) -cne [Convert]::ToHexString([Text.Encoding]::UTF8.GetBytes($expected))) {
        throw 'CSV bytes differ from the expected UTF-8 output.'
    }
    $csvHash = (Get-FileHash -LiteralPath $outputCsv).Hash

    $PSNativeCommandUseErrorActionPreference = $false
    & $exe export-csv --input-path $inputSta --output-path $outputCsv --backend Library 2>$null | Out-Null
    if ($LASTEXITCODE -eq 0) { throw 'Export unexpectedly overwrote an existing CSV.' }
    & $exe export-csv --input-path $inputSta --output-path $outputCsv --backend Library --first-row 9 --force 2>$null | Out-Null
    if ($LASTEXITCODE -eq 0) { throw 'Export unexpectedly accepted an out-of-range row.' }
    $PSNativeCommandUseErrorActionPreference = $true
    if ((Get-FileHash -LiteralPath $outputCsv).Hash -ne $csvHash) { throw 'A failed export changed the original CSV.' }

    & $exe export-csv --input-path $inputSta --output-path $outputCsv --backend Application --columns '3,2,1,2' --first-row 2 --last-row 3 --display-text --force | Out-Null
    $rows = @(Import-Csv -LiteralPath $outputCsv)
    if ($rows.Count -ne 2 -or ($rows[0].PSObject.Properties.Name -join ',') -cne 'case_index,v3,v2,v1' -or
        $rows[0].case_index -ne '2' -or $rows[1].case_index -ne '3' -or $rows[0].v3 -cne 'café "γ"' -or $rows[0].v1 -cne '') {
        throw 'Display text, column order, rows or missing values changed.'
    }

    if ((Get-FileHash -LiteralPath $inputSta).Hash -ne $original) { throw 'Export changed the source spreadsheet.' }
    if (Get-ChildItem -LiteralPath $work -Filter '.statistica-*' -Force) { throw 'A staging file survived cleanup.' }
    'Inspect, CSV bytes, overwrite, rollback and source integrity passed.'
} finally {
    foreach ($file in [IO.Directory]::EnumerateFiles($work)) { [IO.File]::Delete($file) }
    [IO.Directory]::Delete($work)
}
