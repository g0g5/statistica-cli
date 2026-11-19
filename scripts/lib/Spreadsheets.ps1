$staInvariant = [Globalization.CultureInfo]::InvariantCulture

function Read-StaSpec([string]$Path) {
    $spec = Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json -AsHashtable -DateKind String
    if ($spec -isnot [System.Collections.IDictionary]) { throw 'The specification must be a JSON object.' }
    $spec
}

function Get-StaSpreadsheetInfo($Session, [string]$InputPath, [string]$ReportedPath) {
    $sheet = Open-StaDocument $Session 'Spreadsheets' $InputPath
    try {
        Invoke-StaCall 'Read' 'Spreadsheet variable metadata' {
            $columns = for ($index = 1; $index -le $sheet.NumberOfVariables; $index++) {
                [ordered]@{
                    index = $index
                    name = $sheet.VariableName($index)
                    type = $sheet.VariableType($index)
                    missing_code = $sheet.VariableMissingData($index)
                }
            }
            [ordered]@{
                path = $ReportedPath
                cases = $sheet.NumberOfCases
                variables = $sheet.NumberOfVariables
                columns = @($columns)
            }
        }
    } finally { Close-StaDocument $sheet }
}

function Test-StaMissingValue($Value, $Code) {
    $null -eq $Value -or ($Value -is [ValueType] -and $Value -eq $Code)
}

function Export-StaCsv($Session, $Options, [string]$Stage) {
    $sheet = Open-StaDocument $Session 'Spreadsheets' $Options.InputPath
    try {
        $cases = $sheet.NumberOfCases
        $variables = $sheet.NumberOfVariables
        $last = if ($Options.LastRow -eq 0) { $cases } else { [int]$Options.LastRow }
        if ($Options.FirstRow -gt $last -or $last -gt $cases) {
            throw "Row range must be within 1..$cases."
        }
        if ($variables -lt 1) { throw 'The spreadsheet has no variables to export.' }
        $indices = if ($Options['Columns']) {
            foreach ($column in $Options.Columns.Split(',')) {
                $index = 0
                if (-not [int]::TryParse($column, [ref]$index) -or $index -lt 1 -or $index -gt $variables) {
                    throw "Column '$column' is outside 1..$variables."
                }
                $index
            }
        } else { 1..$variables }
        $indices = @($indices | Select-Object -Unique)
        $codes = @{}
        $positions = @{}
        $ranges = [Collections.Generic.List[object]]::new()
        foreach ($index in ($indices | Sort-Object)) {
            $codes[$index] = $sheet.VariableMissingData($index)
            # On ANSI spreadsheets the bulk API can decode native text differently from Cells.Value.
            if ($Options['DisplayText'] -or $sheet.VariableType($index) -eq 1) { continue }
            if ($ranges.Count -eq 0 -or $ranges[-1].Start + $ranges[-1].Count -ne $index) {
                $ranges.Add(@{ Start = $index; Count = 0; Data = $null })
            }
            $range = $ranges[-1]
            $positions[$index] = @{ Range = $range; Offset = $range.Count }
            $range.Count++
        }

        $blockRows = [Math]::Max(1, [Math]::Min(1024, [int][Math]::Floor(65536 / $indices.Count)))
        Invoke-StaCall 'Export' 'Spreadsheet data to CSV' {
            & {
                for ($start = [int]$Options.FirstRow; $start -le $last; $start += $blockRows) {
                    $count = [Math]::Min($blockRows, $last - $start + 1)
                    if (-not $Options['DisplayText']) {
                        foreach ($range in $ranges) {
                            $range.Data = Invoke-StaCall 'Read' 'Spreadsheet.GetData' {
                                ,($sheet.GetData($start, $count, $range.Start, $range.Count, $false))
                            }
                            $range.RowBase = $range.Data.GetLowerBound(0)
                            $range.ColumnBase = $range.Data.GetLowerBound(1)
                        }
                    }
                    for ($offset = 0; $offset -lt $count; $offset++) {
                        $row = $start + $offset
                        $record = [ordered]@{ case_index = $row }
                        foreach ($index in $indices) {
                            if (-not $positions.ContainsKey($index)) {
                                $cell = $sheet.Cells($row, $index)
                                try {
                                    $value = $cell.Value
                                    if (Test-StaMissingValue $value $codes[$index]) { $value = '' }
                                    elseif ($Options['DisplayText']) { $value = $cell.Text }
                                } finally { Release-StaCom $cell }
                            } else {
                                $position = $positions[$index]
                                $range = $position.Range
                                $value = $range.Data.GetValue($range.RowBase + $offset, $range.ColumnBase + $position.Offset)
                                if (Test-StaMissingValue $value $codes[$index]) { $value = '' }
                                elseif ($value -is [double] -or $value -is [single]) { $value = $value.ToString('R', $staInvariant) }
                            }
                            $record["v$index"] = $value
                        }
                        [pscustomobject]$record
                    }
                }
            } | Export-Csv -LiteralPath $Stage -NoTypeInformation -Encoding utf8
        }
    } finally { Close-StaDocument $sheet }
}

function ConvertTo-StaColumn($Entry) {
    if ($Entry -isnot [System.Collections.IDictionary] -or $Entry['name'] -isnot [string] -or
        [string]::IsNullOrWhiteSpace($Entry['name'])) { throw 'Each column requires a name.' }
    foreach ($key in $Entry.Keys) {
        if ($key -cnotin 'name', 'type', 'length', 'formula') { throw "Unknown column property '$key'." }
    }
    $length = $Entry['length']
    switch ($Entry['type']) {
        double {
            if ($null -ne $length) { throw 'Length applies to text columns.' }
            $type = 0
            $length = 0
        }
        text {
            $type = 1
            if (($length -isnot [int] -and $length -isnot [long]) -or $length -lt 1 -or $length -gt [int]::MaxValue) {
                throw 'A text column requires a positive length.'
            }
        }
        default { throw "Column '$($Entry.name)' requires type double or text." }
    }
    [ordered]@{ Name = $Entry.name; Type = $type; Length = [int]$length; Formula = $Entry['formula'] }
}

function Add-StaColumn($Sheet, $Column) {
    Invoke-StaCall 'Write' 'Spreadsheet.AddVariables' {
        $index = $Sheet.NumberOfVariables + 1
        $Sheet.AddVariables($Column.Name, $index - 1, 1, $Column.Type, $Column.Length)
        if ($Sheet.VariableName($index) -cne $Column.Name -or $Sheet.VariableType($index) -ne $Column.Type) {
            throw "STATISTICA could not add the requested column '$($Column.Name)'."
        }
    }
}

function Get-StaColumnStorage($Sheet, [int]$Index) {
    $type = [int]$Sheet.VariableType($Index)
    [ordered]@{
        Type = $type
        Length = if ($type -eq 1) { [int]$Sheet.VariableTypeLength($Index) } else { 0 }
        Code = if ($type -eq 1) { '' } else { [double]$Sheet.VariableMissingData($Index) }
    }
}

function ConvertTo-StaCellValue($Value, $Storage, [string]$Address, [switch]$Csv) {
    if ($null -eq $Value) { return $Storage.Code }
    if ($Storage.Type -eq 1) {
        if ($Value -isnot [string] -or ($Storage.Length -gt 0 -and $Value.Length -gt $Storage.Length)) {
            throw "$Address requires text within the column's length."
        }
        return $Value
    }

    if ($Csv) {
        $number = 0.0
        if (-not [double]::TryParse($Value, [Globalization.NumberStyles]::Float, $staInvariant, [ref]$number)) {
            throw "$Address is not an invariant-culture number."
        }
    } else {
        if ($Value -isnot [ValueType] -or $Value -is [bool]) { throw "$Address requires a JSON number or null." }
        $number = [double]$Value
    }
    if (-not [double]::IsFinite($number) -or $number -eq $Storage.Code -or
        ($Storage.Type -in 2, 3 -and $number -ne [Math]::Truncate($number)) -or
        ($Storage.Type -eq 2 -and ($number -lt [int]::MinValue -or $number -gt [int]::MaxValue)) -or
        ($Storage.Type -eq 3 -and ($number -lt 0 -or $number -gt 255))) {
        throw "$Address is outside the column's numeric range or equals its missing-value code. Use null for a missing value."
    }
    $number
}

function Assert-StaAddress($Sheet, $Row, $Column) {
    if (($Row -isnot [int] -and $Row -isnot [long]) -or ($Column -isnot [int] -and $Column -isnot [long]) -or
        $Row -lt 1 -or $Row -gt $Sheet.NumberOfCases -or $Column -lt 1 -or $Column -gt $Sheet.NumberOfVariables) {
        throw "Spreadsheet cell ($Row,$Column) is out of range."
    }
}

function Set-StaFormula($Sheet, [int]$Index, $Formula) {
    if ($Formula -isnot [string] -or $Formula.Length -lt 2 -or -not $Formula.StartsWith('=')) {
        throw "Column $Index requires a formula starting with '='."
    }
    Invoke-StaCall 'Write' 'Spreadsheet.IsValidFormula' {
        $message = $variables = $first = $last = $null
        if (-not $Sheet.IsValidFormula($Formula, [ref]$message, [ref]$variables, [ref]$first, [ref]$last)) {
            throw "Invalid formula for column $Index. $message"
        }
    }
    Invoke-StaCall 'Write' 'Spreadsheet.VariableLongName and Recalculate' {
        [void]$Sheet.GetType().InvokeMember('VariableLongName', [Reflection.BindingFlags]::SetProperty,
            $null, $Sheet, [object[]]@($Index, $Formula))
        $Sheet.Recalculate($Index)
    }
}

function Get-StaFormulaValues($Sheet, [int]$Index) {
    Invoke-StaCall 'Verify' 'Spreadsheet.GetData (formula)' {
        ,($Sheet.GetData(1, $Sheet.NumberOfCases, $Index, 1, $false))
    }
}

function Confirm-StaCell($Sheet, [int]$Row, [int]$Column, $Expected, [double]$Tolerance) {
    Assert-StaAddress $Sheet $Row $Column
    if ($null -eq $Expected) {
        $matches = $Sheet.MissingData($Row, $Column)
    } elseif ($Expected -is [string]) {
        $matches = $Sheet.Text($Row, $Column) -ceq $Expected
        if (-not $matches) {
            throw "Text at cell ($Row,$Column) did not round-trip. Check the column length and the spreadsheet's character encoding."
        }
    } elseif ($Expected -is [ValueType] -and $Expected -isnot [bool]) {
        $matches = -not $Sheet.MissingData($Row, $Column) -and
            [Math]::Abs($Sheet.Value($Row, $Column) - [double]$Expected) -le $Tolerance
    } else {
        throw "Verification value for ($Row,$Column) must be a number, text, or null."
    }
    if (-not $matches) { throw "Cell ($Row,$Column) differs from its expected value '$Expected'." }
}

function Edit-StaSpreadsheet($Session, $Options, [string]$Stage) {
    $spec = Invoke-StaCall 'Validate' 'EditSta specification' { Read-StaSpec $Options.SpecPath }
    foreach ($key in $spec.Keys) {
        if ($key -cnotin 'addColumns', 'addRows', 'cells', 'formulas') { throw "Unknown EditSta property '$key'." }
    }
    foreach ($key in 'addColumns', 'cells', 'formulas') {
        if ($spec.Contains($key) -and $spec[$key] -isnot [array]) { throw "$key must be an array." }
    }
    $addRows = $spec['addRows'] ?? 0
    if (($addRows -isnot [int] -and $addRows -isnot [long]) -or $addRows -lt 0 -or $addRows -gt [int]::MaxValue) {
        throw 'addRows must be a nonnegative integer.'
    }
    if (-not ($spec['addColumns'] -or $spec['cells'] -or $spec['formulas'] -or $addRows)) {
        throw 'The EditSta specification contains no changes.'
    }

    Copy-StaInput $Options.InputPath $Stage
    $sheet = Open-StaDocument $Session 'Spreadsheets' $Stage -Writable
    $columns = [ordered]@{}
    $formulas = [ordered]@{}
    $written = @{}
    $storageByColumn = @{}
    try {
        Assert-StaWritable $sheet
        if ($addRows) {
            Invoke-StaCall 'Write' 'Spreadsheet.AddCases' {
                $before = $sheet.NumberOfCases
                if ($before + $addRows -gt [int]::MaxValue) { throw 'The requested row count is too large.' }
                $sheet.AddCases($before, [int]$addRows)
                if ($sheet.NumberOfCases -ne $before + $addRows) { throw 'STATISTICA did not add the requested rows.' }
            }
        }
        $names = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        for ($index = 1; $index -le $sheet.NumberOfVariables; $index++) { [void]$names.Add($sheet.VariableName($index)) }
        foreach ($entry in $spec['addColumns']) {
            $column = ConvertTo-StaColumn $entry
            if (-not $names.Add($column.Name)) { throw "The column '$($column.Name)' already exists." }
            Add-StaColumn $sheet $column
            $index = $sheet.NumberOfVariables
            $columns[[string]$index] = $column
            if ($null -ne $column.Formula) { $formulas[[string]$index] = $column.Formula }
        }
        Invoke-StaCall 'Write' 'Spreadsheet.SetData (cell edits)' {
            foreach ($entry in $spec['cells']) {
                if ($entry -isnot [System.Collections.IDictionary] -or -not $entry.Contains('value')) {
                    throw 'Each cell edit requires row, column, and value.'
                }
                $row = $entry['row']
                $index = $entry['column']
                Assert-StaAddress $sheet $row $index
                if (-not $storageByColumn.ContainsKey($index)) {
                    $storageByColumn[$index] = Get-StaColumnStorage $sheet $index
                }
                $address = "$row,$index"
                if ($written.ContainsKey($address)) { throw "Duplicate cell edit ($address)." }
                $data = [object[,]]::new(1, 1)
                $data[0, 0] = ConvertTo-StaCellValue $entry.value $storageByColumn[$index] "Cell ($address)"
                $sheet.SetData([int]$row, [int]$index, $data)
                $written[$address] = if ($entry.value -is [string] -and $entry.value -ceq '') { $null } else { $entry.value }
            }
        }
        foreach ($entry in $spec['formulas']) {
            if ($entry -isnot [System.Collections.IDictionary]) { throw 'Each formula requires column and formula.' }
            Assert-StaAddress $sheet 1 $entry['column']
            $index = [string]$entry.column
            if ($formulas.Contains($index)) { throw "Duplicate formula for column $index." }
            $formulas[$index] = $entry['formula']
        }
        foreach ($index in $formulas.Keys) { Set-StaFormula $sheet ([int]$index) $formulas[$index] }
        $expectedFormulas = @{}
        foreach ($index in $formulas.Keys) { $expectedFormulas[$index] = Get-StaFormulaValues $sheet ([int]$index) }
        $rows = $sheet.NumberOfCases
        $vars = $sheet.NumberOfVariables
        Save-StaSpreadsheet $sheet
    } finally { Close-StaDocument $sheet }
    Assert-StaCleanup

    $sheet = Open-StaDocument $Session 'Spreadsheets' $Stage
    try {
        Invoke-StaCall 'Verify' 'Reopened edited spreadsheet' {
            if ($sheet.NumberOfCases -ne $rows -or $sheet.NumberOfVariables -ne $vars) {
                throw 'Reopened spreadsheet dimensions differ from the saved dimensions.'
            }
            foreach ($index in $columns.Keys) {
                if ($sheet.VariableName([int]$index) -cne $columns[$index].Name -or
                    $sheet.VariableType([int]$index) -ne $columns[$index].Type) {
                    throw "Reopened column $index differs from the saved column."
                }
            }
            foreach ($address in $written.Keys) {
                $row, $index = $address.Split(',')
                if (-not $formulas.Contains($index)) {
                    Confirm-StaCell $sheet ([int]$row) ([int]$index) $written[$address] 0
                }
            }
            foreach ($index in $formulas.Keys) {
                if ($sheet.VariableLongName([int]$index) -cne $formulas[$index]) {
                    throw "Reopened formula differs for column $index."
                }
                $actual = Get-StaFormulaValues $sheet ([int]$index)
                $expected = $expectedFormulas[$index]
                $rowBase = $actual.GetLowerBound(0)
                $columnBase = $actual.GetLowerBound(1)
                for ($r = 0; $r -lt $rows; $r++) {
                    if ($actual.GetValue($rowBase + $r, $columnBase) -cne
                        $expected.GetValue($expected.GetLowerBound(0) + $r, $expected.GetLowerBound(1))) {
                        throw "Reopened formula value differs for column $index, row $($r + 1)."
                    }
                }
            }
        }
    } finally { Close-StaDocument $sheet }
}

function Read-StaCsvRows([string]$Path, $Columns) {
    $parser = [Microsoft.VisualBasic.FileIO.TextFieldParser]::new(
        $Path, [Text.UTF8Encoding]::new($false, $true), $true)
    try {
        $parser.SetDelimiters(',')
        $parser.HasFieldsEnclosedInQuotes = $true
        $parser.TrimWhiteSpace = $false
        $header = $parser.ReadFields()
        if ($null -eq $header -or $header.Count -ne $Columns.Count) { throw 'CSV header does not match the column specification.' }
        for ($column = 0; $column -lt $header.Count; $column++) {
            if ($header[$column] -cne $Columns[$column].Name) {
                throw "CSV header '$($header[$column])' does not match '$($Columns[$column].Name)' in column $($column + 1)."
            }
        }
        while (-not $parser.EndOfData) { ,($parser.ReadFields()) }
    } finally { $parser.Dispose() }
}

function New-StaSpreadsheet($Session, $Options, [string]$Stage, [switch]$Import) {
    $spec = Invoke-StaCall 'Validate' 'Spreadsheet specification' { Read-StaSpec $Options.SpecPath }
    foreach ($key in $spec.Keys) {
        if ($key -cne 'columns' -and ($Import -or $key -cne 'rows')) { throw "Unknown spreadsheet property '$key'." }
    }
    if ($spec['columns'] -isnot [array] -or $spec.columns.Count -lt 1) { throw 'Supply a nonempty columns array.' }
    $columns = @(foreach ($entry in $spec.columns) { ConvertTo-StaColumn $entry })
    if (@($columns.Name | Sort-Object -Unique).Count -ne $columns.Count -or
        ($columns | Where-Object { $null -ne $_.Formula })) {
        throw 'Column names must be unique. Use EditSta to set formulas.'
    }
    if ($Import) {
        $rows = @(Invoke-StaCall 'Read' 'CSV parsing' { Read-StaCsvRows $Options.InputPath $columns })
    } else {
        if ($spec['rows'] -isnot [array]) { throw 'Supply a rows array.' }
        $rows = $spec.rows
    }
    if ($rows.Count -lt 1) { throw 'Supply at least one data row.' }

    $documents = $Session.App.Spreadsheets
    $sheet = $null
    $blocks = [Collections.Generic.List[object]]::new()
    try {
        $sheet = Invoke-StaCall 'Create' 'Spreadsheets.New' { ,($documents.New()) }
        $defaults = $sheet.NumberOfVariables
        foreach ($column in $columns) { Add-StaColumn $sheet $column }
        if ($defaults -gt 0) { $sheet.DeleteVariables(1, $defaults) }
        $cases = $sheet.NumberOfCases
        if ($cases -lt $rows.Count) { $sheet.AddCases($cases, $rows.Count - $cases) }
        if ($cases -gt $rows.Count) { $sheet.DeleteCases($rows.Count + 1, $cases) }

        $storage = @(for ($index = 1; $index -le $columns.Count; $index++) { Get-StaColumnStorage $sheet $index })
        $blockRows = [Math]::Max(1, [Math]::Min(1024, [int][Math]::Floor(65536 / $columns.Count)))
        for ($start = 0; $start -lt $rows.Count; $start += $blockRows) {
            $count = [Math]::Min($blockRows, $rows.Count - $start)
            $data = [object[,]]::new($count, $columns.Count)
            for ($r = 0; $r -lt $count; $r++) {
                $row = $rows[$start + $r]
                if ($row -isnot [array] -or $row.Count -ne $columns.Count) {
                    throw "Data row $($start + $r + 1) must contain $($columns.Count) fields."
                }
                for ($c = 0; $c -lt $columns.Count; $c++) {
                    $value = $row[$c]
                    if ($Import -and ($value -ceq '' -or
                        ($Options.ContainsKey('MissingToken') -and $value -ceq $Options.MissingToken))) { $value = $null }
                    $data[$r, $c] = ConvertTo-StaCellValue $value $storage[$c] "Cell ($($start + $r + 1),$($c + 1))" -Csv:$Import
                }
            }
            Invoke-StaCall 'Write' 'Spreadsheet.SetData' { $sheet.SetData($start + 1, 1, $data) }
            $blocks.Add($data)
        }
        Save-StaDocument $sheet $Stage
    } finally {
        Close-StaDocument $sheet
        Release-StaCom $documents
    }
    Assert-StaCleanup

    $sheet = Open-StaDocument $Session 'Spreadsheets' $Stage
    try {
        Invoke-StaCall 'Verify' 'Reopened new spreadsheet' {
            if ($sheet.NumberOfCases -ne $rows.Count -or $sheet.NumberOfVariables -ne $columns.Count) {
                throw 'Reopened spreadsheet dimensions differ from the specification.'
            }
            for ($index = 1; $index -le $columns.Count; $index++) {
                if ($sheet.VariableName($index) -cne $columns[$index - 1].Name -or
                    $sheet.VariableType($index) -ne $columns[$index - 1].Type) {
                    throw "Reopened column $index differs from the specification."
                }
            }
            $start = 1
            foreach ($data in $blocks) {
                $actual = $sheet.GetData($start, $data.GetLength(0), 1, $data.GetLength(1), $false)
                $rBase = $actual.GetLowerBound(0)
                $cBase = $actual.GetLowerBound(1)
                for ($r = 0; $r -lt $data.GetLength(0); $r++) {
                    for ($c = 0; $c -lt $data.GetLength(1); $c++) {
                        if ($data[$r, $c] -cne $actual.GetValue($r + $rBase, $c + $cBase)) {
                            throw "Reopened cell ($($start + $r),$($c + 1)) differs from the specification."
                        }
                    }
                }
                $start += $data.GetLength(0)
            }
        }
    } finally { Close-StaDocument $sheet }
}
