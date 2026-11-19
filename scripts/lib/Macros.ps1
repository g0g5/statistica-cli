function Invoke-StaMacro($Session, $Options, $Job) {
    $expectations = [Collections.Generic.List[object]]::new()
    if ($Options['VerifyCells']) {
        foreach ($address in $Options.VerifyCells.Keys) {
            if ($address -notmatch '^([1-9]\d*),([1-9]\d*)$') {
                throw "Invalid VerifyCells address '$address'. Use row,column with 1-based indices."
            }
            $row = [int]$Matches[1]
            $column = [int]$Matches[2]
            $value = $Options.VerifyCells[$address]
            if ($null -ne $value -and $value -isnot [string] -and
                ($value -isnot [ValueType] -or $value -is [bool] -or -not [double]::IsFinite($value))) {
                throw "VerifyCells value for ($row,$column) must be a finite number, text, or null."
            }
            $expectations.Add(@{
                Row = $row; Column = $column; Value = $value
            })
        }
    }
    $tolerance = $Options['Tolerance'] ?? 1e-8
    $visible = $Job.NeedsDesktop
    $sheet = $macro = $null
    Copy-StaInput $Options.InputPath $Job.Stage
    try {
        $sheet = Open-StaDocument $Session 'Spreadsheets' $Job.Stage -Writable
        Assert-StaWritable $sheet
        if ($visible) { Invoke-StaCall 'Open' 'Spreadsheet.Visible' { $sheet.Visible = $true } }
        $macro = Open-StaDocument $Session 'Macros' $Options.MacroPath
        if ($visible) {
            Invoke-StaCall 'Open' 'Spreadsheet.Activate' { $sheet.Activate() }
            Invoke-StaCall 'Execute' 'Macro.Execute' { $macro.Execute() }
        } else {
            Invoke-StaCall 'Execute' 'Macro.ExecuteWithArgument' { $macro.ExecuteWithArgument($sheet) }
        }
        Save-StaSpreadsheet $sheet
    } finally {
        $openedMacro = $null -ne $macro
        Close-StaDocument $macro -LeaveForQuit:$visible
        Close-StaDocument $sheet -LeaveForQuit:$visible
        if ($openedMacro) { Clear-StaMacroDocuments $Session -LeaveForQuit:$visible }
    }
    if ($visible) { Exit-StaSession }
    Assert-StaCleanup

    # Use the original backend for verification, including after desktop Quit.
    $verifySession = Get-StaSession $Job $Session.Backend
    $sheet = Open-StaDocument $verifySession 'Spreadsheets' $Job.Stage
    try {
        Invoke-StaCall 'Verify' 'Reopened macro output' {
            [void]$sheet.NumberOfCases
            [void]$sheet.NumberOfVariables
            foreach ($cell in $expectations) {
                Confirm-StaCell $sheet $cell.Row $cell.Column $cell.Value $tolerance
            }
        }
    } finally { Close-StaDocument $sheet }
}
