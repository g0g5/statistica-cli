# STATISTICA Command-Line Guide

Run the published `statistica-cli.exe` on Windows with a locally licensed STATISTICA installation. The examples use PowerShell and assume `statistica-cli` is on PATH. The CLI launches a separate STA worker for STATISTICA commands. Use the examples with your own documents.

```powershell
$source = 'data/measurements.sta'
$work = 'analysis'
New-Item -ItemType Directory -Path $work -Force | Out-Null

statistica-cli probe --detailed
statistica-cli inspect --input-path $source
statistica-cli export-csv --input-path $source --output-path "$work/preview.csv" --columns '1,2' --last-row 10
Import-Csv -LiteralPath "$work/preview.csv" | Select-Object -First 2
```

## Commands

| Command | Main options and output |
| --- | --- |
| `probe` | JSON with selected COM component and desktop activation state. `--detailed` adds available component registrations, PIDs, worker architecture and apartment. In Auto mode it checks both COM components. |
| `inspect` | `--input-path` for an STA. JSON with dimensions, original variable names, type codes, and missing-value codes. |
| `export-csv` | `--input-path` and `--output-path`. Optional `--columns '1,3,7'`, `--first-row`, `--last-row`, `--display-text`. |
| `run-macro` | `--input-path`, `--macro-path`, and `--output-path` for a new STA. Use `--pass-input-as-script-argument` for macros that read `GetScriptArgument`. Optional `--verify-cells` and `--tolerance`. |
| `edit-sta` | `--input-path`, `--spec-path`, and `--output-path`. Apply row additions, columns, cell values, and formulas to an STA copy. |
| `create-sta` | `--spec-path` and `--output-path`. Create a typed STA from JSON rows. |
| `import-csv` | `--input-path`, `--spec-path`, and `--output-path`. Import UTF-8 CSV into a typed STA. Optional `--missing-token`. |
| `create-graph` | `--input-path` for an STA, `--graph-type Line` or `Scatter`, `--y-columns`, `--output-path` for an STG. Scatter needs `--x-column` and supports `--connect-points`. |
| `export-graph` | `--input-path` for an STG and `--output-path` for PNG or SVG. |
| `screenshot` | `--input-path` for an STA, STG, or STW and `--output-path` for PNG. Optional `--crop-to-document` and `--capture-timeout-seconds`. |
| `list-workbook` | `--input-path` for an STW. JSON with item indices, names, paths, and types. |
| `extract-workbook-item` | `--input-path`, `--output-path`, and one of `--item-index` or `--item-path`. Extract a spreadsheet or graph to its native file format. |
| `build-workbook` | `--document-paths` for the ordered list of STA/STG files and `--output-path` for a new STW. |
| `install-skill` | Install this skill and guide offline. Optional `--global` installs in the user profile, and `--force` replaces conflicting files. |
| `batch` | `--manifest-path` for a JSON job list. Returns per-job JSON records. Optional `--continue-on-error`. |

Successful file-producing commands return the absolute output path. Output directories must exist. Outputs use fresh names by default. Pass `--force` to replace an existing output after the new file is saved and verified. Writing commands work on separate paths from their inputs.

CSV exports use UTF-8 and include `case_index` and headers such as `v1`, preserving the requested order. Missing values become empty fields. Use `--display-text` for the text shown by STATISTICA, including text labels, formatting control characters, and locale-specific number formats. The default exports internal codes for numeric variables with text labels.

## Edit and create spreadsheets

`edit-sta` accepts a JSON specification. Rows and columns are 1-based. New columns are appended in their listed order. This example appends a column that contains the difference between the first two numeric columns.

```json
{
  "addColumns": [
    {"name": "Check", "type": "double", "formula": "=V1-V2"}
  ]
}
```

```powershell
statistica-cli edit-sta --input-path $source --spec-path "$work/check.json" --output-path "$work/checked.sta"
statistica-cli export-csv --input-path "$work/checked.sta" --output-path "$work/check.csv"
```

A specification may also include `"addRows": 2`, `"cells": [{"row": 2, "column": 1, "value": 0}]`, and `"formulas": [{"column": 3, "formula": "=V1-V2"}]`. Formula columns are recalculated in declaration order after cell edits. Use either the added column's `formula` or the `formulas` list for a given column. Numeric values are JSON numbers, text values are strings, and `null` means missing. Old STA formats can have a limited character set. Text that changes during save and read-back fails verification.

Use `create-sta` with a `columns` and `rows` specification. Each column requires `name` and `type`. Text columns also need a positive `length`.

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
statistica-cli create-sta --spec-path "$work/parameters.json" --output-path "$work/parameters_cli.sta"
statistica-cli import-csv --input-path "$work/parameters.csv" --spec-path "$work/columns.json" --missing-token NA --output-path "$work/imported_cli.sta"
```

For `import-csv`, put only the `columns` array in the specification. The CSV header must match names and order exactly. Numeric fields use a decimal point and may use exponent notation. Empty fields and the optional exact missing token become missing values. Quoted commas, quotes and multiline fields are parsed as CSV.

## Run macros and batches

By default, `run-macro` opens a temporary copy and activates it in the desktop Application. For a self-contained SVB that uses `GetScriptArgument`, pass the spreadsheet explicitly. Use `--verify-cells` with a JSON file containing expected numeric, text, or missing values. The standalone EXE also accepts an inline JSON object. Numeric comparisons use an absolute tolerance, defaulting to `1e-8`. Expected strings are compared with the display text.

To adapt a macro that uses `Set sheet = ActiveSpreadsheet`, make a copy and use `Set sheet = GetScriptArgument()` in that copy. [TIBCO Macro object and script arguments](https://docs.tibco.com/pub/stat/14.0.0/doc/html/StatisticaObjectModelHelp/GUID-F9A500EF-DE69-4C34-9D1F-05639AEA3B2F.html)

```powershell
'{"1,3":12.5,"2,3":null}' | Set-Content "$work/expected.json"
statistica-cli run-macro --input-path $source --macro-path 'macros/calculate.svb' --pass-input-as-script-argument --backend Library --output-path "$work/calculated.sta" --verify-cells "$work/expected.json"
```

Manifest entries use the PascalCase field names shown below. Direct calls use lowercase, hyphenated options. Relative paths resolve from the manifest directory. In this example, place `measurements.sta` and `check.json` beside `jobs.json`. A job can consume an earlier job's output within the same batch. Hidden jobs reuse a compatible COM session. Visible macro jobs and screenshots each use their own desktop instance.

```json
{
  "jobs": [
    {
      "Action": "EditSta",
      "InputPath": "measurements.sta",
      "SpecPath": "check.json",
      "OutputPath": "batch-checked.sta"
    },
    {
      "Action": "ExportCsv",
      "InputPath": "batch-checked.sta",
      "OutputPath": "batch-checked.csv",
      "Columns": "1,2,3"
    }
  ]
}
```

```powershell
statistica-cli batch --manifest-path "$work/jobs.json" --diagnostic-path "$work/jobs.jsonl"
```

Batch summaries contain `succeeded`, `failed`, `skipped`, or `timed_out` statuses, with results or output paths for successful jobs. A failure stops the batch by default. `--continue-on-error` runs independent later jobs after a regular job error. Failed dependencies are skipped. If a COM call blocks, the worker stops and completed, verified jobs can still be published. A batch with incomplete jobs reports an error after writing its JSON summary.

## Create charts, capture windows, and assemble workbooks

```powershell
statistica-cli create-graph --input-path "$work/checked.sta" --graph-type Line --y-columns 1 2 3 --output-path "$work/trends.stg"
statistica-cli create-graph --input-path "$work/checked.sta" --graph-type Scatter --x-column 1 --y-columns 2 --connect-points --output-path "$work/scatter.stg"
statistica-cli export-graph --input-path "$work/trends.stg" --output-path "$work/trends.svg"
statistica-cli screenshot --input-path "$work/trends.stg" --output-path "$work/trends-window.png"

statistica-cli build-workbook --document-paths "$work/checked.sta" "$work/trends.stg" --output-path "$work/results.stw"
statistica-cli list-workbook --input-path "$work/results.stw"
statistica-cli extract-workbook-item --input-path "$work/results.stw" --item-index 2 --output-path "$work/extracted-graph.stg"
```

Use numeric X and Y columns for scatterplots. Line plots use case order on the X axis. Multiple Y columns produce one graph with multiple series. Scatterplots have no fitted curve by default, and `--connect-points` enables lines between points. To export PNG instead of SVG, change the output extension. [TIBCO graph object model](https://docs.tibco.com/pub/stat/14.0.0/doc/html/StatisticaObjectModelHelp/GUID-24C29661-0DDF-4D94-AD96-CA7CFA8239B5.html)

Screenshots need an unlocked interactive Windows desktop and the Application backend. By default they include the STATISTICA main window and its menus. `--crop-to-document` keeps the target document area. The CLI checks the active document and waits for two stable, nonblank samples. The default capture window is eight seconds. `export-graph` works with a hidden session.

The workbook list can contain duplicate paths. Pass an item index to select an occurrence unambiguously. When assembling a workbook, top-level item names are taken from the input filenames. [TIBCO workbook traversal](https://docs.tibco.com/pub/stat/14.0.0/doc/html/StatisticaObjectModelHelp/GUID-86A6727B-DF85-46EC-A337-4C492EAF347A.html)

## Diagnose desktop COM activation

`--backend Auto` tries `Application` first for standalone `probe`, `inspect`, `export-csv`, and `run-macro`. Other hidden commands prefer `Library`. Use `--backend Application` or `--backend Library` to select one explicitly. `screenshot` and `run-macro` in active-spreadsheet mode require the desktop backend. Fallback occurs during activation, before any macro execution. [TIBCO Application and Library object model](https://docs.tibco.com/pub/stat/14.0.0/doc/html/StatisticaObjectModelHelp/GUID-273F4553-1310-4340-AE28-C95864185A9D.html)

```powershell
statistica-cli probe --detailed --verbose --diagnostic-path "$work/activation_cli.jsonl"
Get-CimInstance Win32_Process -Filter "Name = 'statist.exe'" |
    Select-Object ProcessId, SessionId, CreationDate, CommandLine
```

`--diagnostic-path` creates a JSON Lines log in an existing directory. It records operation boundaries, timing, errors, and process identity when COM reports it. `--timeout-seconds` defaults to 120 seconds per job. On timeout the CLI terminates its worker. It may also terminate an out-of-process server when the COM-reported PID, creation time, session, and process command line establish that the server belongs to that worker. If activation hangs before returning an identity, the log records process candidates for investigation.

For SVB macros that open their own inputs and save their own results, STATISTICA 14.0 also supports `statist.exe /RunMacro=<absolute macro path>`. Record an analysis macro in STATISTICA to discover the analysis object and its dialog options before automating a new analysis. [TIBCO Quick Reference Guide](https://docs.tibco.com/pub/stat/14.0.0/doc/pdf/TIB_stat_14.0_quick_ref.pdf)

## Install the skill

`statistica-cli install-skill` writes `SKILL.md` and `references/cli-guide.md` to `.agents/skills/statistica-cli/` under the current directory. Add `--global` to use the current user's home directory. The files are embedded in the executable and install offline. Reinstalling the same version succeeds. Use `--force` to replace conflicting files.

## Run the integration checks

```powershell
./tests/smoke.ps1
./tests/smoke.ps1 -IncludeDesktop
```

Publish the Native AOT executable before running these commands from the CLI repository root. The smoke test reads its inputs from `tests/fixtures` and writes outputs to a temporary directory. Use `-IncludeDesktop` in an unlocked desktop session to add checks for the active-spreadsheet macro, window capture, and cleanup after a timed-out desktop COM call.
