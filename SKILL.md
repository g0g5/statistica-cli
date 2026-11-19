---
name: statistica-cli
description: Operate TIBCO STATISTICA from pwsh on Windows for sm3ABP coursework. Diagnose COM, read and write native spreadsheets, run SVB macros, generate and capture charts, and work with STATISTICA workbooks.
---

# STATISTICA CLI

Run `D:/Projects/Coding/statistical-cli/scripts/statistica.ps1` from the coursework repository root with pwsh 7.5 or newer. The CLI uses the locally licensed STATISTICA COM interface. Start with `Probe -Detailed` and see the [command-line guide](references/cli-guide.md) for options and examples.

- Use the course materials in the corresponding lab directory. Save results to a working directory or a temporary directory. STATISTICA cells use 1-based row and column indices. Use `Inspect` to check the actual column positions in older STA files.
- Use `RunMacro` for an SVB that expects `ActiveSpreadsheet`. For a macro that reads `GetScriptArgument`, add `-PassInputAsScriptArgument` to allow the headless Library backend. Read the macro to confirm its inputs and any additional documents it opens or writes.
- Use `Batch` for repeated work on independent copies. Use `EditSta` for verification columns and formulas and `CreateSta` or `ImportCsv` for explicitly typed tables.
- Use `CreateGraph` and `ExportGraph` for native figures. Use `Screenshot` in an unlocked desktop session when the evidence needs to show STATISTICA's window. Save the result before capturing it.
- Use `ListWorkbook` to obtain item indices, then `ExtractWorkbookItem` or `BuildWorkbook` to work with embedded spreadsheets and graphs.
- On a COM failure, use `-Verbose` and `-DiagnosticPath` to record the phase, HRESULT, process, and timing. Verify results through `Inspect`, `ExportCsv`, or `RunMacro -VerifyCells`. See `lab6/PLAN.md` in the coursework repository for reference values.
