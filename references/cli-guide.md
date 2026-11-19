# STATISTICA Command-Line Guide

Run the CLI from pwsh 7.5 or newer on Windows with a locally licensed STATISTICA installation. The CLI launches a separate STA worker for each invocation. Run the examples from the sm3ABP coursework repository root.

```powershell
$cli = 'D:/Projects/Coding/statistical-cli/scripts/statistica.ps1'
$source = 'lab6/materials/PR_DWLS_025_005_065_20.sta'
$macro = 'lab6/materials/Trd_DWR.svb'
$work = 'lab6/work'
New-Item -ItemType Directory -Path $work -Force | Out-Null

& $cli Probe -Detailed
& $cli Inspect -InputPath $source
& $cli RunMacro -InputPath $source -MacroPath $macro -OutputPath "$work/profit_dwls_cli.sta"
& $cli ExportCsv -InputPath "$work/profit_dwls_cli.sta" -OutputPath "$work/trend_cli.csv" -Columns '1,2,32,33,34,35' -FirstRow 1 -LastRow 20
Import-Csv -LiteralPath "$work/trend_cli.csv" | Select-Object -First 2
```

In row 1, `v33` should be approximately `207.332991` and `v34` approximately `-44.735329`. In rows 1 and 2, `v35` should be approximately `122.060906` and `1.770847`, respectively.

## Commands

| Command | Main options and output |
| --- | --- |
| `Probe` | JSON with selected COM component and desktop activation state. `-Detailed` adds available component registrations, PIDs, worker architecture and apartment. In Auto mode it checks both COM components. |
| `Inspect` | `-InputPath` for an STA. JSON with dimensions, original variable names, type codes, and missing-value codes. |
| `ExportCsv` | `-InputPath` and `-OutputPath`. Optional `-Columns '1,3,7'`, `-FirstRow`, `-LastRow`, `-DisplayText`. |
| `RunMacro` | `-InputPath`, `-MacroPath`, and `-OutputPath` for a new STA. Use `-PassInputAsScriptArgument` for macros that read `GetScriptArgument`. Optional `-VerifyCells` and `-Tolerance`. |
| `EditSta` | `-InputPath`, `-SpecPath`, and `-OutputPath`. Apply row additions, columns, cell values, and formulas to an STA copy. |
| `CreateSta` | `-SpecPath` and `-OutputPath`. Create a typed STA from JSON rows. |
| `ImportCsv` | `-InputPath`, `-SpecPath`, and `-OutputPath`. Import UTF-8 CSV into a typed STA. Optional `-MissingToken`. |
| `CreateGraph` | `-InputPath` for an STA, `-GraphType Line` or `Scatter`, `-YColumns`, `-OutputPath` for an STG. Scatter needs `-XColumn` and supports `-ConnectPoints`. |
| `ExportGraph` | `-InputPath` for an STG and `-OutputPath` for PNG or SVG. |
| `Screenshot` | `-InputPath` for an STA, STG, or STW and `-OutputPath` for PNG. Optional `-CropToDocument` and `-CaptureTimeoutSeconds`. |
| `ListWorkbook` | `-InputPath` for an STW. JSON with item indices, names, paths, and types. |
| `ExtractWorkbookItem` | `-InputPath`, `-OutputPath`, and one of `-ItemIndex` or `-ItemPath`. Extract a spreadsheet or graph to its native file format. |
| `BuildWorkbook` | `-DocumentPaths` for the ordered list of STA/STG files and `-OutputPath` for a new STW. |
| `Batch` | `-ManifestPath` for a JSON job list. Returns per-job JSON records. Optional `-ContinueOnError`. |

Successful file-producing commands return the absolute output path. Output directories must exist. Outputs use fresh names by default. Pass `-Force` to replace an existing output after the new file is saved and verified. Writing commands work on separate paths from their inputs.

CSV exports use UTF-8 and include `case_index` and headers such as `v1`, preserving the requested order. Missing values become empty fields. Use `-DisplayText` for the text shown by STATISTICA, including text labels, formatting control characters, and locale-specific number formats. The default exports internal codes for numeric variables with text labels.

## Edit and create spreadsheets

`EditSta` accepts a JSON specification. Rows and columns are 1-based. New columns are appended in their listed order. Apply the example to the Lab 6 trend result and verify that `Check` is zero within floating-point error.

```json
{
  "addColumns": [
    {"name": "Check", "type": "double", "formula": "=V1-V33-V34"}
  ]
}
```

```powershell
& $cli EditSta -InputPath "$work/profit_dwls_cli.sta" -SpecPath "$work/check.json" -OutputPath "$work/profit_checked_cli.sta"
& $cli ExportCsv -InputPath "$work/profit_checked_cli.sta" -OutputPath "$work/check_cli.csv" -Columns '38'
```

A specification may also include `"addRows": 2`, `"cells": [{"row": 48, "column": 3, "value": 0}]`, and `"formulas": [{"column": 38, "formula": "=V1-V33-V34"}]`. Formula columns are recalculated in declaration order after cell edits. Use either the added column's `formula` or the `formulas` list for a given column. Numeric values are JSON numbers, text values are strings, and `null` means missing. Old STA formats can have a limited character set. Text that changes during save and read-back fails verification.

Use `CreateSta` with a `columns` and `rows` specification. Each column requires `name` and `type`. Text columns also need a positive `length`.

```json
{
  "columns": [
    {"name": "Parameter", "type": "double"},
    {"name": "Comment", "type": "text", "length": 64}
  ],
  "rows": [[0.25, "first pass"], [0.30, null]]
}
```

```powershell
& $cli CreateSta -SpecPath "$work/parameters.json" -OutputPath "$work/parameters_cli.sta"
& $cli ImportCsv -InputPath "$work/parameters.csv" -SpecPath "$work/columns.json" -MissingToken NA -OutputPath "$work/imported_cli.sta"
```

For `ImportCsv`, put only the `columns` array in the specification. The CSV header must match names and order exactly. Numeric fields use a decimal point and may use exponent notation. Empty fields and the optional exact missing token become missing values. Quoted commas, quotes and multiline fields are parsed as CSV.

## Run macros and batches

By default, `RunMacro` opens a temporary copy and activates it in the desktop Application. For a self-contained SVB that uses `GetScriptArgument`, pass the spreadsheet explicitly. Use `-VerifyCells` to check numeric, text, or missing values after reopening the saved STA. Numeric comparisons use an absolute tolerance, defaulting to `1e-8`. Expected strings are compared with the display text.

For Lab 7, make a copy of `Trd_Lws.svb` and replace `Set TS = ActiveSpreadsheet` with `Set TS = GetScriptArgument()` in the copy. [TIBCO Macro object and script arguments](https://docs.tibco.com/pub/stat/14.0.0/doc/html/StatisticaObjectModelHelp/GUID-F9A500EF-DE69-4C34-9D1F-05639AEA3B2F.html)

```powershell
$source = 'lab7/materials/Lowess/Trd/PR_Lws_0425_005_080_11_3.sta'
$work = 'lab7/work'
& $cli RunMacro -InputPath $source -MacroPath "$work/trend_argument.svb" -PassInputAsScriptArgument -Backend Library -OutputPath "$work/lowess_cli.sta" -VerifyCells @{'1,33'=210.827438527348; '42,3'=0}
```

Manifest entries use the same option names as direct calls. Relative paths resolve from the manifest directory. A job can consume an earlier job's output within the same batch. Hidden jobs reuse a compatible COM session. Visible macro jobs and screenshots each use their own desktop instance.

```json
{
  "jobs": [
    {
      "Action": "EditSta",
      "InputPath": "profit_dwls_cli.sta",
      "SpecPath": "check.json",
      "OutputPath": "checked_cli.sta"
    },
    {
      "Action": "ExportCsv",
      "InputPath": "checked_cli.sta",
      "OutputPath": "checked_cli.csv",
      "Columns": "1,33,34,38"
    }
  ]
}
```

```powershell
$work = 'lab6/work'
& $cli Batch -ManifestPath "$work/jobs.json" -DiagnosticPath "$work/jobs.jsonl"
```

Batch summaries contain `succeeded`, `failed`, `skipped`, or `timed_out` statuses, with results or output paths for successful jobs. A failure stops the batch by default. `-ContinueOnError` runs independent later jobs after a regular job error. Failed dependencies are skipped. If a COM call blocks, the worker stops and completed, verified jobs can still be published. A batch with incomplete jobs reports an error after writing its JSON summary.

## Create charts, capture windows, and assemble workbooks

```powershell
$work = 'lab6/work'
& $cli CreateGraph -InputPath "$work/profit_checked_cli.sta" -GraphType Line -YColumns 1,2,33 -OutputPath "$work/trends_cli.stg"
& $cli CreateGraph -InputPath "$work/profit_checked_cli.sta" -GraphType Scatter -XColumn 1 -YColumns 2 -ConnectPoints -OutputPath "$work/profit_scatter_cli.stg"
& $cli ExportGraph -InputPath "$work/trends_cli.stg" -OutputPath "$work/trends_cli.svg"
& $cli Screenshot -InputPath "$work/trends_cli.stg" -OutputPath "$work/trends_window_cli.png"

& $cli ListWorkbook -InputPath 'lab4/lab4.stw'
& $cli ExtractWorkbookItem -InputPath 'lab4/lab4.stw' -ItemIndex 3 -OutputPath "$work/lab4_graph_cli.stg"
& $cli BuildWorkbook -DocumentPaths @("$work/profit_checked_cli.sta", "$work/trends_cli.stg") -OutputPath "$work/results_cli.stw"
```

Use numeric X and Y columns for scatterplots. For example, compare convergence iterations or numeric parameter values with DRF. Line plots use case order on the X axis. Multiple Y columns produce one graph with multiple series. Scatterplots have no fitted curve by default, and `-ConnectPoints` enables lines between points. To export PNG instead of SVG, change the output extension. [TIBCO graph object model](https://docs.tibco.com/pub/stat/14.0.0/doc/html/StatisticaObjectModelHelp/GUID-24C29661-0DDF-4D94-AD96-CA7CFA8239B5.html)

Screenshots need an unlocked interactive Windows desktop and the Application backend. By default they include the STATISTICA main window and its menus. `-CropToDocument` keeps the target document area. The CLI checks the active document and waits for two stable, nonblank samples. The default capture window is eight seconds. `ExportGraph` works with a hidden session.

The workbook list can contain duplicate paths. Pass an item index to select an occurrence unambiguously. When assembling a workbook, top-level item names are taken from the input filenames. [TIBCO workbook traversal](https://docs.tibco.com/pub/stat/14.0.0/doc/html/StatisticaObjectModelHelp/GUID-86A6727B-DF85-46EC-A337-4C492EAF347A.html)

## Diagnose desktop COM activation

`-Backend Auto` tries `Application` first for standalone `Probe`, `Inspect`, `ExportCsv`, and `RunMacro`. Other hidden commands prefer `Library`. Use `-Backend Application` or `-Backend Library` to select one explicitly. `Screenshot` and `RunMacro` in active-spreadsheet mode require the desktop backend. Fallback occurs during activation, before any macro execution. [TIBCO Application and Library object model](https://docs.tibco.com/pub/stat/14.0.0/doc/html/StatisticaObjectModelHelp/GUID-273F4553-1310-4340-AE28-C95864185A9D.html)

```powershell
& $cli Probe -Detailed -Verbose -DiagnosticPath "$work/activation_cli.jsonl"
Get-CimInstance Win32_Process -Filter "Name = 'statist.exe'" |
    Select-Object ProcessId, SessionId, CreationDate, CommandLine
```

`-DiagnosticPath` creates a JSON Lines log in an existing directory. It records operation boundaries, timing, errors, and process identity when COM reports it. `-TimeoutSeconds` defaults to 120 seconds per job. On timeout the CLI terminates its worker. It may also terminate an out-of-process server when the COM-reported PID, creation time, session, and process command line establish that the server belongs to that worker. If activation hangs before returning an identity, the log records process candidates for investigation. [Repository issue #1](https://github.com/g0g5/IU5M-sm3ABP/issues/1)

For SVB macros that open their own inputs and save their own results, STATISTICA 14.0 also supports `statist.exe /RunMacro=<absolute macro path>`. Record an analysis macro in STATISTICA to discover the analysis object and its dialog options before automating a new analysis. [TIBCO Quick Reference Guide](https://docs.tibco.com/pub/stat/14.0.0/doc/pdf/TIB_stat_14.0_quick_ref.pdf)

## Run the integration checks

```powershell
./tests/smoke.ps1
./tests/smoke.ps1 -IncludeDesktop
```

Run these commands from the CLI repository root. The smoke test reads its inputs from `tests/fixtures` and writes outputs to a temporary directory. Use `-IncludeDesktop` in an unlocked desktop session to add checks for the active-spreadsheet macro, window capture, and cleanup after a timed-out desktop COM call.
