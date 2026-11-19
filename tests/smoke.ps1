#Requires -Version 7.5
[CmdletBinding()]
param(
    [string]$Executable = (Join-Path $PSScriptRoot '../src/statistica-cli/bin/Release/net11.0/win-x64/publish/statistica-cli.exe'),
    [switch]$IncludeDesktop
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$cli = (Resolve-Path -LiteralPath $Executable).Path
$PSNativeCommandUseErrorActionPreference = $true
$fixtures = Join-Path $PSScriptRoot 'fixtures'
$lab3 = Join-Path $fixtures 'lab3.csv'
$book = Join-Path $fixtures 'lab4.stw'
$lab6 = Join-Path $fixtures 'PR_DWLS_025_005_065_20.sta'
$lab7 = Join-Path $fixtures 'PR_Lws_0425_005_080_11_3.sta'
$macro = Join-Path $fixtures 'Trd_Lws.svb'
$inputs = @($lab3, $book, $lab6, $lab7, $macro)
$before = (Get-FileHash -LiteralPath $inputs).Hash
$baseline = @(Get-Process -Name statist -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)

function Assert-Result([bool]$Condition, [string]$Message) {
    if (-not $Condition) { throw $Message }
}

function Write-TestJson([string]$Name, $Value) {
    ConvertTo-Json -InputObject $Value -Depth 12 | Set-Content -LiteralPath (Join-Path $temp $Name) -Encoding utf8
}

$temp = Join-Path ([IO.Path]::GetTempPath()) ('statistica-cli-smoke-' + [guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($temp)
try {
    $probeBackend = if ($IncludeDesktop) { 'Auto' } else { 'Library' }
    $probe = & $cli probe --backend $probeBackend --detailed | ConvertFrom-Json
    Assert-Result ($probe.components.Library.available -and $probe.host.apartment -eq 'STA') 'Library STA probe failed.'
    if ($IncludeDesktop) { Assert-Result $probe.components.Application.available 'Application probe failed.' }

    $info = & $cli inspect --input-path $lab6 --backend Library | ConvertFrom-Json
    Assert-Result ($info.cases -eq 48 -and $info.variables -eq 37) 'Lab 6 dimensions differ.'
    $csv = Join-Path $temp 'legacy.csv'
    $null = & $cli export-csv --input-path $lab6 --output-path $csv --columns '37,35,2,35' --last-row 3 --backend Library
    $rows = Import-Csv -LiteralPath $csv
    Assert-Result (($rows[0].PSObject.Properties.Name -join ',') -eq 'case_index,v37,v35,v2') 'CSV columns changed.'
    Assert-Result ($rows[0].v37 -eq '102' -and $rows[2].v35 -eq '') 'CSV raw values or missing values changed.'
    $display = Join-Path $temp 'display.csv'
    $null = & $cli export-csv --input-path $lab6 --output-path $display --columns '37' --last-row 2 --display-text --backend Library
    Assert-Result ((Import-Csv -LiteralPath $display)[1].v37 -eq 'Feb-2008') 'CSV display labels changed.'

    Write-TestJson 'edit.json' @{ addColumns = @(@{ name = 'Check'; type = 'double'; formula = '=V1-V33-V34' }) }
    $edited = Join-Path $temp '检查 Данные.sta'
    $null = & $cli edit-sta --input-path $lab6 --spec-path (Join-Path $temp 'edit.json') --output-path $edited
    $checkCsv = Join-Path $temp 'check.csv'
    $null = & $cli export-csv --input-path $edited --output-path $checkCsv --columns '38' --backend Library
    $check = Import-Csv -LiteralPath $checkCsv
    $different = @($check | Where-Object { $_.v38 -eq '' -or [Math]::Abs([double]$_.v38) -gt 1e-9 })
    Assert-Result ($check.Count -eq 48 -and $different.Count -eq 0) 'Lab 6 formula differs.'

    Write-TestJson 'ansi-edit.json' @{
        addColumns = @(@{ name = 'AnsiText'; type = 'text'; length = 32 })
        cells = @(@{ row = 1; column = 38; value = 'Москва' })
    }
    $ansiSta = Join-Path $temp 'legacy-ansi.sta'
    $ansiCsv = Join-Path $temp 'legacy-ansi.csv'
    $null = & $cli edit-sta --input-path $lab6 --spec-path (Join-Path $temp 'ansi-edit.json') --output-path $ansiSta --backend Library
    $ansiBytes = [Text.Encoding]::GetEncoding(1251).GetBytes('Москва')
    $foundAnsi = [Text.Encoding]::Latin1.GetString([IO.File]::ReadAllBytes($ansiSta)).Contains([Text.Encoding]::Latin1.GetString($ansiBytes))
    Assert-Result $foundAnsi 'The legacy STA did not store the Cyrillic label as ANSI.'
    $null = & $cli export-csv --input-path $ansiSta --output-path $ansiCsv --columns 38 --last-row 2 --display-text --backend Library
    $ansiRows = @(Import-Csv -LiteralPath $ansiCsv)
    Assert-Result ($ansiRows[0].v38 -ceq 'Москва' -and $ansiRows[1].v38 -ceq '') 'Legacy ANSI display text or missing value changed.'

    $schema = @(@{ name = 'Parameter'; type = 'double' }, @{ name = 'Comment'; type = 'text'; length = 100 })
    Write-TestJson 'create.json' @{ columns = $schema; rows = @(@(0.25, '北京, Москва'), @($null, $null)) }
    Write-TestJson 'columns.json' @{ columns = $schema }
    $created = Join-Path $temp 'typed.sta'
    $null = & $cli create-sta --spec-path (Join-Path $temp 'create.json') --output-path $created
    Write-TestJson 'formula-chain.json' @{
        addRows = 1
        addColumns = @(@{ name = 'First'; type = 'double'; formula = '=V1*2' }, @{ name = 'Second'; type = 'double' })
        cells = @(@{ row = 3; column = 1; value = 4 }, @{ row = 3; column = 2; value = 'edited, text' })
        formulas = @(@{ column = 4; formula = '=V3+3' })
    }
    $formulaSta = Join-Path $temp 'formula-chain.sta'
    $formulaCsv = Join-Path $temp 'formula-chain.csv'
    $null = & $cli edit-sta --input-path $created --spec-path (Join-Path $temp 'formula-chain.json') --output-path $formulaSta
    $null = & $cli export-csv --input-path $formulaSta --output-path $formulaCsv --first-row 3
    $formulaRow = Import-Csv -LiteralPath $formulaCsv
    Assert-Result ($formulaRow.v2 -ceq 'edited, text' -and $formulaRow.v3 -eq '8' -and $formulaRow.v4 -eq '11') 'Edits, new rows or formula order changed.'

    $csvInput = Join-Path $temp 'import.csv'
    @('Parameter,Comment', '0.25,"北京, Москва"', 'NA,"2026-11-19T02:00:00Z"') |
        Set-Content -LiteralPath $csvInput -Encoding utf8BOM
    $imported = Join-Path $temp 'imported.sta'
    $null = & $cli import-csv --input-path $csvInput --spec-path (Join-Path $temp 'columns.json') --missing-token NA --output-path $imported
    $roundTrip = Join-Path $temp 'import roundtrip.csv'
    $null = & $cli export-csv --input-path $imported --output-path $roundTrip --display-text --backend Library
    $importRows = Import-Csv -LiteralPath $roundTrip
    Assert-Result ($importRows[0].v2 -ceq '北京, Москва' -and $importRows[1].v1 -eq '' -and
        $importRows[1].v2 -ceq '2026-11-19T02:00:00Z') 'Typed CSV import did not round-trip.'

    $lab3Columns = @((Get-Content -LiteralPath $lab3 -TotalCount 1).Split(',') | ForEach-Object {
        @{ name = $_; type = 'double' }
    })
    Write-TestJson 'series-columns.json' @{ columns = $lab3Columns }
    $series = Join-Path $temp 'series.sta'
    $null = & $cli import-csv --input-path $lab3 --spec-path (Join-Path $temp 'series-columns.json') --output-path $series --backend Library
    $seriesCsv = Join-Path $temp 'series.csv'
    $null = & $cli export-csv --input-path $series --output-path $seriesCsv --columns '1,2,9' --backend Library
    $seriesRows = @(Import-Csv -LiteralPath $seriesCsv)
    Assert-Result ($seriesRows.Count -eq 16512) 'Lab 3 round-trip row count changed.'
    Assert-Result ($seriesRows[0].v2 -eq '112' -and $seriesRows[0].v9 -eq '' -and
        $seriesRows[143].v2 -eq '432' -and $seriesRows[144].v2 -eq '' -and
        $seriesRows[155].v1 -eq '156') 'Lab 3 series values or missing fields changed.'
    $nonemptyTail = @($seriesRows | Select-Object -Skip 156 | Where-Object { $_.v1 -ne '' -or $_.v2 -ne '' -or $_.v9 -ne '' })
    Assert-Result ($nonemptyTail.Count -eq 0) 'Lab 3 all-missing tail did not round-trip.'

    $expected = @{ '1,33' = 210.827438527348; '42,3' = 0 }
    $encoding = [Text.Encoding]::Latin1
    $argumentMacro = Join-Path $temp 'LOWESS argument.svb'
    [IO.File]::WriteAllText($argumentMacro, [regex]::Replace([IO.File]::ReadAllText($macro, $encoding),
        '(?im)^(\s*Set\s+TS\s*=\s*)ActiveSpreadsheet\s*$', '$1GetScriptArgument()'), $encoding)
    Write-TestJson 'jobs.json' @{ jobs = @(
        @{ Action = 'RunMacro'; InputPath = $lab7; MacroPath = 'LOWESS argument.svb'; OutputPath = 'lowess.sta'
            PassInputAsScriptArgument = $true; VerifyCells = $expected }
        @{ Action = 'CreateGraph'; InputPath = 'lowess.sta'; OutputPath = 'trend.stg'; GraphType = 'Line'; YColumns = @(1, 2, 33) }
        @{ Action = 'CreateGraph'; InputPath = 'lowess.sta'; OutputPath = 'scatter.stg'; GraphType = 'Scatter'
            XColumn = 1; YColumns = @(2, 33); ConnectPoints = $true }
        @{ Action = 'ExportGraph'; InputPath = 'trend.stg'; OutputPath = 'trend.png' }
        @{ Action = 'ExportGraph'; InputPath = 'scatter.stg'; OutputPath = 'scatter.svg' }
        @{ Action = 'BuildWorkbook'; DocumentPaths = @('lowess.sta', 'trend.stg'); OutputPath = 'assembled.stw' }
        @{ Action = 'ListWorkbook'; InputPath = 'assembled.stw' }
    ) }
    $batchLog = Join-Path $temp 'batch.jsonl'
    $PSNativeCommandUseErrorActionPreference = $false
    $batch = & $cli batch --manifest-path (Join-Path $temp 'jobs.json') --backend Library --diagnostic-path $batchLog | ConvertFrom-Json
    $incomplete = @($batch.jobs | Where-Object status -ne 'succeeded')
    Assert-Result ($LASTEXITCODE -eq 0 -and $incomplete.Count -eq 0) ($incomplete | ConvertTo-Json -Depth 6 -Compress)
    $PSNativeCommandUseErrorActionPreference = $true
    $starts = @(Get-Content -LiteralPath $batchLog | ForEach-Object { $_ | ConvertFrom-Json } | Where-Object event -eq 'server')
    Assert-Result ($starts.Count -eq 1 -and $batch.jobs[-1].result.items[1].name -eq 'trend') 'Batch session or item name changed.'

    $items = (& $cli list-workbook --input-path $book | ConvertFrom-Json).items
    foreach ($type in 'Spreadsheet', 'Graph') {
        $item = $items | Where-Object type -eq $type | Select-Object -First 1
        $extension = if ($type -eq 'Spreadsheet') { '.sta' } else { '.stg' }
        $null = & $cli extract-workbook-item --input-path $book --item-index $item.index --output-path (Join-Path $temp "embedded$extension")
    }

    $duplicateBook = Join-Path $temp 'duplicates.stw'
    $null = & $cli build-workbook --document-paths $created $created --output-path $duplicateBook
    $duplicate = & $cli list-workbook --input-path $duplicateBook | ConvertFrom-Json
    Assert-Result ($duplicate.items.Count -eq 2 -and $duplicate.items[0].path -ceq $duplicate.items[1].path) 'Duplicate workbook paths changed.'
    $PSNativeCommandUseErrorActionPreference = $false
    $ambiguous = Join-Path $temp 'ambiguous.sta'
    $message = (& $cli extract-workbook-item --input-path $duplicateBook --item-path $duplicate.items[0].path --output-path $ambiguous 2>&1 | Out-String)
    Assert-Result ($LASTEXITCODE -ne 0 -and $message -like '*ambiguous*1, 2*' -and -not [IO.File]::Exists($ambiguous)) 'An ambiguous workbook path was accepted.'
    $PSNativeCommandUseErrorActionPreference = $true
    $null = & $cli extract-workbook-item --input-path $duplicateBook --item-index 2 --output-path (Join-Path $temp 'selected.sta')
    $assembled = Join-Path $temp 'assembled.stw'
    $null = & $cli extract-workbook-item --input-path $assembled --item-path $batch.jobs[-1].result.items[1].path --output-path (Join-Path $temp 'selected.stg')

    $racePath = Join-Path $temp 'race.csv'
    foreach ($force in $false, $true) {
        $commands = @(foreach ($lastRow in 1024, 2048) {
            $info = [Diagnostics.ProcessStartInfo]::new($cli)
            $info.UseShellExecute = $false
            $info.CreateNoWindow = $true
            $info.RedirectStandardOutput = $true
            $info.RedirectStandardError = $true
            foreach ($argument in @('export-csv', '--input-path', $series, '--columns', '1,2,9',
                    '--last-row', [string]$lastRow, '--output-path', $racePath, '--backend', 'Library')) {
                $info.ArgumentList.Add($argument)
            }
            if ($force) { $info.ArgumentList.Add('--force') }
            [Diagnostics.Process]::Start($info)
        })
        try {
            foreach ($process in $commands) { $process.WaitForExit() }
            $successful = @($commands | Where-Object ExitCode -eq 0)
            $expectedWriters = if ($force) { 2 } else { 1 }
            $errors = ($commands | ForEach-Object { $_.StandardError.ReadToEnd() }) -join ' '
            Assert-Result ($successful.Count -eq $expectedWriters) "Unexpected concurrent publish outcome. $errors"
            $rowCount = @(Import-Csv -LiteralPath $racePath).Count
            Assert-Result ($rowCount -in 1024, 2048) 'Concurrent publication produced an incomplete CSV.'
        } finally {
            foreach ($process in $commands) { $process.Dispose() }
        }
    }

    $existingHash = (Get-FileHash -LiteralPath $created).Hash
    $lock = [IO.File]::Open($created, 'Open', 'Read', 'Read')
    try {
        $PSNativeCommandUseErrorActionPreference = $false
        $publishError = (& $cli create-sta --spec-path (Join-Path $temp 'create.json') --output-path $created --force 2>&1 | Out-String).Trim()
        Assert-Result ($LASTEXITCODE -ne 0 -and $publishError -like '[[]Publish/File.Move]*') 'A locked output did not fail at publish.'
        $PSNativeCommandUseErrorActionPreference = $true
    } finally { $lock.Dispose() }
    Assert-Result ((Get-FileHash -LiteralPath $created).Hash -eq $existingHash) 'A failed publish replaced the original file.'

    'Sub Main', '    Err.Raise 5', 'End Sub' | Set-Content -LiteralPath (Join-Path $temp 'fail.svb') -Encoding ascii
    Write-TestJson 'failing jobs.json' @{ jobs = @(
        @{ Action = 'RunMacro'; InputPath = $lab7; MacroPath = 'fail.svb'; PassInputAsScriptArgument = $true; OutputPath = 'failed.sta' }
        @{ Action = 'ExportCsv'; InputPath = 'failed.sta'; OutputPath = 'skipped.csv' }
        @{ Action = 'Inspect'; InputPath = $lab6 }
    ) }
    $PSNativeCommandUseErrorActionPreference = $false
    $failedBatch = & $cli batch --manifest-path (Join-Path $temp 'failing jobs.json') --backend Library --continue-on-error 2>$null |
        ConvertFrom-Json
    Assert-Result ($LASTEXITCODE -ne 0 -and ($failedBatch.jobs.status -join ',') -eq 'failed,skipped,succeeded' -and
        $failedBatch.jobs[0].error -like '[[]Execute/Macro.ExecuteWithArgument*') 'Batch error did not preserve dependency and phase.'
    $PSNativeCommandUseErrorActionPreference = $true

    if ($IncludeDesktop) {
        Write-TestJson 'macro expectations.json' $expected
        $null = & $cli run-macro --input-path $lab7 --macro-path $macro --output-path (Join-Path $temp 'active.sta') `
            --verify-cells (Join-Path $temp 'macro expectations.json')
        $null = & $cli screenshot --input-path (Join-Path $temp 'trend.stg') --output-path (Join-Path $temp 'window.png')
        $null = & $cli screenshot --input-path $edited --crop-to-document --output-path (Join-Path $temp 'crop.png')
        $null = & $cli screenshot --input-path $assembled --output-path (Join-Path $temp 'book-window.png')
        $null = & $cli screenshot --input-path $edited --output-path (Join-Path $temp 'full-sheet.png')
        $full = [IO.File]::ReadAllBytes((Join-Path $temp 'full-sheet.png'))
        $crop = [IO.File]::ReadAllBytes((Join-Path $temp 'crop.png'))
        $fullWidth = [Net.IPAddress]::NetworkToHostOrder([BitConverter]::ToInt32($full, 16))
        $fullHeight = [Net.IPAddress]::NetworkToHostOrder([BitConverter]::ToInt32($full, 20))
        $cropWidth = [Net.IPAddress]::NetworkToHostOrder([BitConverter]::ToInt32($crop, 16))
        $cropHeight = [Net.IPAddress]::NetworkToHostOrder([BitConverter]::ToInt32($crop, 20))
        Assert-Result ($cropWidth -le $fullWidth -and $cropHeight -lt $fullHeight) 'Document crop does not fit inside the main window.'

        $slow = Join-Path $temp 'slow.svb'
        'Sub Main', '    Dim started As Double', '    started = Timer', '    Do While Timer - started < 25',
            '    Loop', 'End Sub' | Set-Content -LiteralPath $slow -Encoding ascii
        Write-TestJson 'timed jobs.json' @{ jobs = @(
            @{ Action = 'ExportCsv'; InputPath = $lab6; OutputPath = 'before-timeout.csv'; LastRow = 2 }
            @{ Action = 'RunMacro'; InputPath = $lab7; MacroPath = 'slow.svb'; PassInputAsScriptArgument = $true; OutputPath = 'timeout.sta' }
            @{ Action = 'Inspect'; InputPath = 'timeout.sta' }
        ) }
        $timeoutLog = Join-Path $temp 'timeout.jsonl'
        $PSNativeCommandUseErrorActionPreference = $false
        $timed = & $cli batch --manifest-path (Join-Path $temp 'timed jobs.json') --backend Application --continue-on-error `
            --timeout-seconds 8 --diagnostic-path $timeoutLog 2>$null | ConvertFrom-Json
        Assert-Result ($LASTEXITCODE -ne 0 -and ($timed.jobs.status -join ',') -eq 'succeeded,timed_out,skipped' -and
            $timed.jobs[1].error -like '[[]Execute/Macro.ExecuteWithArgument]*exceeded*') 'Timeout did not preserve checkpoints and identify the blocked call.'
        $PSNativeCommandUseErrorActionPreference = $true
        Assert-Result ([IO.File]::Exists((Join-Path $temp 'before-timeout.csv'))) 'A verified output before a timeout was lost.'
        Assert-Result (-not [IO.File]::Exists((Join-Path $temp 'timeout.sta'))) 'A timed-out macro published an output.'
        $log = Get-Content -LiteralPath $timeoutLog | ForEach-Object { $_ | ConvertFrom-Json }
        $desktop = $log | Where-Object event -eq 'server' | Select-Object -First 1
        Assert-Result ($desktop.pid -and -not (Get-Process -Id $desktop.pid -ErrorAction SilentlyContinue)) 'Desktop server remained after timeout.'
    }

    Assert-Result (-not (Get-ChildItem -LiteralPath $temp -Force -Filter '.statistica-*')) 'A stage was left behind.'
    Assert-Result (-not (Compare-Object $before (Get-FileHash -LiteralPath $inputs).Hash)) 'Test fixtures changed.'
    $remaining = @(Get-Process -Name statist -ErrorAction SilentlyContinue | Where-Object { $_.Id -notin $baseline })
    Assert-Result ($remaining.Count -eq 0) 'STATISTICA servers remained after the smoke test.'
    'STATISTICA CLI smoke checks passed.'
} finally {
    $resolved = (Resolve-Path -LiteralPath $temp).ProviderPath
    if ([IO.Path]::GetDirectoryName($resolved) -eq [IO.Path]::GetTempPath().TrimEnd('\') -and
        [IO.Path]::GetFileName($resolved) -match '^statistica-cli-smoke-[0-9a-f]{32}$') {
        Remove-Item -LiteralPath $resolved -Recurse -Force
    }
}
