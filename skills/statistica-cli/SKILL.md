---
name: statistica-cli
description: Automate TIBCO STATISTICA on Windows. Inspect, create, edit, import and export spreadsheets, run SVB macros, generate and capture graphs, and manage workbook items.
---

# STATISTICA CLI

Run `statistica-cli` on Windows with a locally licensed STATISTICA installation. Start with `statistica-cli probe --detailed` to check the available COM backends. See the [command-line guide](references/cli-guide.md) for options and examples.

- Use `inspect` to check STA dimensions and column indices. Spreadsheet rows and columns are 1-based. Use `edit-sta` for cell edits or formula columns, and `create-sta` or `import-csv` to build typed tables.
- Use `run-macro` in an unlocked desktop session when an SVB expects `ActiveSpreadsheet`. For an SVB that reads `GetScriptArgument()`, add `--pass-input-as-script-argument`. Read the macro to identify its inputs and any other documents it opens or writes.
- Use `create-graph` and `export-graph` for native graphs and image exports. Use `screenshot` in an unlocked desktop session to capture the STATISTICA window or document area.
- Use `list-workbook` to find item indices, then `extract-workbook-item` or `build-workbook` to work with embedded spreadsheets and graphs.
- Use `batch` for ordered operations and dependencies. Resolve manifest paths relative to the manifest and give outputs separate paths from their inputs.
- On a COM failure, use `--verbose` and `--diagnostic-path` to see the operation, HRESULT, process identity, and timing. Verify results with `inspect`, `export-csv`, or `run-macro --verify-cells`.
