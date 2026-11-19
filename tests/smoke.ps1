#Requires -Version 7.5
[CmdletBinding()]
param([switch]$IncludeDesktop)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$cli = Join-Path $PSScriptRoot '../scripts/statistica.ps1'
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
    $probe = & $cli Probe -Backend $probeBackend -Detailed | ConvertFrom-Json
    Assert-Result ($probe.components.Library.available -and $probe.host.apartment -eq 'STA') 'Library STA probe failed.'
    if ($IncludeDesktop) { Assert-Result $probe.components.Application.available 'Application probe failed.' }

    $info = & $cli Inspect -InputPath $lab6 -Backend Library | ConvertFrom-Json
    Assert-Result ($info.cases -eq 48 -and $info.variables -eq 37) 'Lab 6 dimensions differ.'
    $csv = Join-Path $temp 'legacy.csv'
    $null = & $cli ExportCsv -InputPath $lab6 -OutputPath $csv -Columns '37,35,2,35' -LastRow 3 -Backend Library
    $rows = Import-Csv -LiteralPath $csv
    Assert-Result (($rows[0].PSObject.Properties.Name -join ',') -eq 'case_index,v37,v35,v2') 'CSV columns changed.'
    Assert-Result ($rows[0].v37 -eq '102' -and $rows[2].v35 -eq '') 'CSV raw values or missing values changed.'
    $display = Join-Path $temp 'display.csv'
    $null = & $cli ExportCsv -InputPath $lab6 -OutputPath $display -Columns '37' -LastRow 2 -DisplayText -Backend Library
    Assert-Result ((Import-Csv -LiteralPath $display)[1].v37 -eq 'Feb-2008') 'CSV display labels changed.'

    Write-TestJson 'edit.json' @{ addColumns = @(@{ name = 'Check'; type = 'double'; formula = '=V1-V33-V34' }) }
    $edited = Join-Path $temp '检查 Данные.sta'
    $null = & $cli EditSta -InputPath $lab6 -SpecPath (Join-Path $temp 'edit.json') -OutputPath $edited
    $checkCsv = Join-Path $temp 'check.csv'
    $null = & $cli ExportCsv -InputPath $edited -OutputPath $checkCsv -Columns '38' -Backend Library
    $check = Import-Csv -LiteralPath $checkCsv
    $different = @($check | Where-Object { $_.v38 -eq '' -or [Math]::Abs([double]$_.v38) -gt 1e-9 })
    Assert-Result ($check.Count -eq 48 -and $different.Count -eq 0) 'Lab 6 formula differs.'

    $schema = @(@{ name = 'Parameter'; type = 'double' }, @{ name = 'Comment'; type = 'text'; length = 100 })
    Write-TestJson 'create.json' @{ columns = $schema; rows = @(@(0.25, '北京, Москва'), @($null, $null)) }
    Write-TestJson 'columns.json' @{ columns = $schema }
    $created = Join-Path $temp 'typed.sta'
    $null = & $cli CreateSta -SpecPath (Join-Path $temp 'create.json') -OutputPath $created
    $csvInput = Join-Path $temp 'import.csv'
    @('Parameter,Comment', '0.25,"北京, Москва"', 'NA,"2026-11-19T02:00:00Z"') |
        Set-Content -LiteralPath $csvInput -Encoding utf8BOM
    $imported = Join-Path $temp 'imported.sta'
    $null = & $cli ImportCsv -InputPath $csvInput -SpecPath (Join-Path $temp 'columns.json') -MissingToken NA -OutputPath $imported
    $roundTrip = Join-Path $temp 'import roundtrip.csv'
    $null = & $cli ExportCsv -InputPath $imported -OutputPath $roundTrip -DisplayText -Backend Library
    $importRows = Import-Csv -LiteralPath $roundTrip
    Assert-Result ($importRows[0].v2 -ceq '北京, Москва' -and $importRows[1].v1 -eq '' -and
        $importRows[1].v2 -ceq '2026-11-19T02:00:00Z') 'Typed CSV import did not round-trip.'

    $lab3Columns = @((Get-Content -LiteralPath $lab3 -TotalCount 1).Split(',') | ForEach-Object {
        @{ name = $_; type = 'double' }
    })
    Write-TestJson 'series-columns.json' @{ columns = $lab3Columns }
    $series = Join-Path $temp 'series.sta'
    $null = & $cli ImportCsv -InputPath $lab3 -SpecPath (Join-Path $temp 'series-columns.json') -OutputPath $series -Backend Library
    $seriesCsv = Join-Path $temp 'series.csv'
    $null = & $cli ExportCsv -InputPath $series -OutputPath $seriesCsv -Columns '1,2,9' -Backend Library
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
    $batch = & $cli Batch -ManifestPath (Join-Path $temp 'jobs.json') -Backend Library -DiagnosticPath $batchLog | ConvertFrom-Json
    Assert-Result (@($batch.jobs | Where-Object status -ne 'succeeded').Count -eq 0) 'Batch did not complete.'
    $starts = @(Get-Content -LiteralPath $batchLog | ForEach-Object { $_ | ConvertFrom-Json } | Where-Object event -eq 'server')
    Assert-Result ($starts.Count -eq 1 -and $batch.jobs[-1].result.items[1].name -eq 'trend') 'Batch session or item name changed.'

    $items = (& $cli ListWorkbook -InputPath $book | ConvertFrom-Json).items
    foreach ($type in 'Spreadsheet', 'Graph') {
        $item = $items | Where-Object type -eq $type | Select-Object -First 1
        $extension = if ($type -eq 'Spreadsheet') { '.sta' } else { '.stg' }
        $null = & $cli ExtractWorkbookItem -InputPath $book -ItemIndex $item.index -OutputPath (Join-Path $temp "embedded$extension")
    }

    $existingHash = (Get-FileHash -LiteralPath $created).Hash
    $lock = [IO.File]::Open($created, 'Open', 'Read', 'Read')
    try {
        $publishError = $null
        try { $null = & $cli CreateSta -SpecPath (Join-Path $temp 'create.json') -OutputPath $created -Force }
        catch { $publishError = $_.Exception.Message }
        Assert-Result ($publishError -like '[[]Publish/File.Move]*') 'A locked output did not fail at publish.'
    } finally { $lock.Dispose() }
    Assert-Result ((Get-FileHash -LiteralPath $created).Hash -eq $existingHash) 'A failed publish replaced the original file.'

    'Sub Main', '    Err.Raise 5', 'End Sub' | Set-Content -LiteralPath (Join-Path $temp 'fail.svb') -Encoding ascii
    Write-TestJson 'failing jobs.json' @{ jobs = @(
        @{ Action = 'RunMacro'; InputPath = $lab7; MacroPath = 'fail.svb'; PassInputAsScriptArgument = $true; OutputPath = 'failed.sta' }
        @{ Action = 'ExportCsv'; InputPath = 'failed.sta'; OutputPath = 'skipped.csv' }
        @{ Action = 'Inspect'; InputPath = $lab6 }
    ) }
    $summary = [Collections.Generic.List[string]]::new()
    $failure = $null
    try {
        & $cli Batch -ManifestPath (Join-Path $temp 'failing jobs.json') -Backend Library -ContinueOnError |
            ForEach-Object { $summary.Add($_) }
    } catch { $failure = $_.Exception.Message }
    $failedBatch = $summary[0] | ConvertFrom-Json
    Assert-Result ($failure -and ($failedBatch.jobs.status -join ',') -eq 'failed,skipped,succeeded' -and
        $failedBatch.jobs[0].error -like '[[]Execute/Macro.ExecuteWithArgument*') 'Batch error did not preserve dependency and phase.'

    if ($IncludeDesktop) {
        $null = & $cli RunMacro -InputPath $lab7 -MacroPath $macro -OutputPath (Join-Path $temp 'active.sta') -VerifyCells $expected
        $null = & $cli Screenshot -InputPath (Join-Path $temp 'trend.stg') -OutputPath (Join-Path $temp 'window.png')
        $null = & $cli Screenshot -InputPath $edited -CropToDocument -OutputPath (Join-Path $temp 'crop.png')
        $slow = Join-Path $temp 'slow.svb'
        'Sub Main', '    Dim started As Double', '    started = Timer', '    Do While Timer - started < 25',
            '    Loop', 'End Sub' | Set-Content -LiteralPath $slow -Encoding ascii
        $timeoutOptions = @{
            Action = 'RunMacro'; InputPath = $lab7; MacroPath = $slow; PassInputAsScriptArgument = $true; Backend = 'Application'
            OutputPath = Join-Path $temp 'timeout.sta'; TimeoutSeconds = 8; DiagnosticPath = Join-Path $temp 'timeout.jsonl'
        }
        $timedOut = $null
        try { $null = & $cli @timeoutOptions }
        catch { $timedOut = $_.Exception.Message }
        Assert-Result ($timedOut -like '[[]Execute/Macro.ExecuteWithArgument]*exceeded*') 'Timeout did not identify the blocked call.'
        Assert-Result (-not [IO.File]::Exists($timeoutOptions.OutputPath)) 'A timed-out macro published an output.'
        $log = Get-Content -LiteralPath $timeoutOptions.DiagnosticPath | ForEach-Object { $_ | ConvertFrom-Json }
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
