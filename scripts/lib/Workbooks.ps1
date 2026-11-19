$staWorkbookTypes = @('Folder', 'Spreadsheet', 'Graph', 'Report', 'Macro', 'Other')

function Get-StaWorkbookEntries($Workbook) {
    $item = $Workbook.Root
    $index = 0
    try {
        while ($null -ne $item) {
            $next = $Workbook.NextItem($item)
            Release-StaCom $item
            $item = $next
            if ($null -eq $item) { break }
            $index++
            $type = [int]$item.Type
            [ordered]@{
                index = $index
                name = $item.Name
                path = $item.ItemPath
                type = if ($type -ge 0 -and $type -lt $staWorkbookTypes.Count) { $staWorkbookTypes[$type] } else { 'Other' }
                type_code = $type
            }
        }
    } finally { Release-StaCom $item }
}

function Get-StaWorkbookInfo($Session, [string]$InputPath, [string]$ReportedPath) {
    $book = Open-StaDocument $Session 'Workbooks' $InputPath
    try {
        $entries = @(Invoke-StaCall 'Read' 'Workbook.NextItem' { Get-StaWorkbookEntries $book })
        [ordered]@{ path = $ReportedPath; items = $entries }
    } finally { Close-StaDocument $book }
}

function Export-StaWorkbookItem($Session, $Options, [string]$Stage) {
    $book = Open-StaDocument $Session 'Workbooks' $Options.InputPath
    $doc = $node = $null
    try {
        $entries = @(Invoke-StaCall 'Read' 'Workbook.NextItem' { Get-StaWorkbookEntries $book })
        if ($Options.ContainsKey('ItemIndex')) { $selected = @($entries | Where-Object index -eq $Options.ItemIndex) }
        else { $selected = @($entries | Where-Object { $_.path -ceq $Options.ItemPath }) }
        if ($selected.Count -gt 1) {
            throw "ItemPath is ambiguous. Use one of these ItemIndex values: $($selected.index -join ', ')."
        }
        if ($selected.Count -eq 0) { throw 'No matching workbook item.' }
        $selected = $selected[0]
        $extension = [IO.Path]::GetExtension($Stage)
        $type = if ($extension -eq '.sta') { 'Spreadsheet' } else { 'Graph' }
        if ($selected.type -ne $type) { throw "Workbook item $($selected.index) is $($selected.type), which cannot be saved as $extension." }

        $node = $book.Root
        for ($index = 0; $index -lt $selected.index; $index++) {
            $next = $book.NextItem($node)
            Release-StaCom $node
            $node = $next
        }
        # scWorkbookExtractCopy preserves the item in the source workbook.
        $doc = Invoke-StaCall 'Extract' 'WorkbookItem.Extract(copy)' { ,($node.Extract(1)) }
        Save-StaDocument $doc $Stage
    } finally {
        Close-StaDocument $doc
        Release-StaCom $node
        Close-StaDocument $book
    }
    Assert-StaCleanup
    $reopened = Open-StaDocument $Session ($type + 's') $Stage
    Close-StaDocument $reopened
}

function New-StaWorkbook($Session, [string[]]$Paths, [string]$Stage, [string[]]$ReportedPaths) {
    $documents = $Session.App.Workbooks
    $book = $root = $null
    $expected = [Collections.Generic.List[object]]::new()
    try {
        $book = Invoke-StaCall 'Create' 'Workbooks.New' { ,($documents.New()) }
        $root = $book.Root
        for ($index = 0; $index -lt $Paths.Count; $index++) {
            $path = $Paths[$index]
            $type = if ([IO.Path]::GetExtension($path) -eq '.sta') { 'Spreadsheet' } else { 'Graph' }
            $name = [IO.Path]::GetFileNameWithoutExtension($ReportedPaths[$index])
            $item = $null
            try {
                $item = Invoke-StaCall 'Write' 'Workbook.InsertFile' { ,($book.InsertFile($path, $root, 1)) }
                Invoke-StaCall 'Write' 'WorkbookItem.Name' { $item.Name = $name }
                $expected.Add(@{ type = $type; name = $name })
            } finally { Release-StaCom $item }
        }
        Save-StaDocument $book $Stage
    } finally {
        Release-StaCom $root
        Close-StaDocument $book
        Release-StaCom $documents
    }
    Assert-StaCleanup

    $book = Open-StaDocument $Session 'Workbooks' $Stage
    try {
        Invoke-StaCall 'Verify' 'Reopened workbook' {
            $entries = @(Get-StaWorkbookEntries $book)
            if ($entries.Count -ne $expected.Count) { throw 'The reopened workbook has a different item count.' }
            for ($index = 0; $index -lt $entries.Count; $index++) {
                if ($entries[$index].type -ne $expected[$index].type -or $entries[$index].name -cne $expected[$index].name) {
                    throw "Reopened workbook item $($index + 1) differs from its input."
                }
            }
        }
    } finally { Close-StaDocument $book }
}
