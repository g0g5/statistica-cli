function Set-StaGraphConnections($Session, $Graph) {
    $path = Join-Path $PSScriptRoot '../svb/SetPlotStyle.svb'
    $macro = Open-StaDocument $Session 'Macros' $path
    try {
        Invoke-StaCall 'Execute' 'Plot2D style macro' { $macro.ExecuteWithArgument($Graph) }
    } finally { Close-StaDocument $macro }
}

function New-StaGraph($Session, $Options, [string]$Stage) {
    $sheet = $analysis = $dialog = $graphs = $graph = $null
    $series = @($Options.YColumns)
    try {
        $sheet = Open-StaDocument $Session 'Spreadsheets' $Options.InputPath
        $variables = $sheet.NumberOfVariables
        foreach ($index in $series + @($Options['XColumn'] | Where-Object { $null -ne $_ })) {
            if ($index -gt $variables) { throw "Graph column $index is outside 1..$variables." }
            if ($sheet.VariableType($index) -eq 1) { throw "Graph column $index must be numeric." }
        }
        $id = if ($Options.GraphType -eq 'Line') { 11012 } else { 11003 }
        $analysis = Invoke-StaCall 'Graph' "Application.Analysis($id)" { ,($Session.App.Analysis($id, $sheet)) }
        $dialog = Invoke-StaCall 'Graph' 'Analysis.Dialog' { ,$analysis.Dialog }
        Invoke-StaCall 'Graph' 'Dialog.GraphType and FitType' {
            # Each dialog uses 1 for its multiple-Y mode and 0 for no fit.
            $dialog.GraphType = if ($series.Count -gt 1) { 1 } else { 0 }
            $dialog.FitType = 0
        }
        $list = $series -join ' '
        if ($Options.GraphType -eq 'Scatter') { $list = "$($Options.XColumn) |$list" }
        Invoke-StaCall 'Graph' 'Dialog.Variables' { $dialog.Variables = $list }
        $graphs = Invoke-StaCall 'Graph' 'Dialog.Graphs' { ,$dialog.Graphs }
        $graph = Invoke-StaCall 'Graph' 'Graphs.Item(1)' {
            if ($graphs.Count -ne 1) { throw "Expected one graph, received $($graphs.Count)." }
            ,($graphs.Item(1))
        }
        if ($graph.PlotsCount -ne $series.Count) { throw "Expected $($series.Count) plots, received $($graph.PlotsCount)." }
        if ($Options['ConnectPoints']) { Set-StaGraphConnections $Session $graph }
        Save-StaDocument $graph $Stage
    } finally {
        Close-StaDocument $graph
        Release-StaCom $graphs
        Release-StaCom $dialog
        if ($analysis) { Invoke-StaCall 'Cleanup' 'Analysis.Close' { $analysis.Close() } -Cleanup }
        Release-StaCom $analysis
        Close-StaDocument $sheet
    }
    Assert-StaCleanup

    $graph = Open-StaDocument $Session 'Graphs' $Stage
    try {
        Invoke-StaCall 'Verify' 'Reopened graph' {
            if ($graph.PlotsCount -ne $series.Count) { throw 'The reopened graph has a different number of plots.' }
        }
    } finally { Close-StaDocument $graph }
}

function Export-StaGraph($Session, $Options, [string]$Stage) {
    $graph = Open-StaDocument $Session 'Graphs' $Options.InputPath
    try {
        switch ([IO.Path]::GetExtension($Stage)) {
            '.png' {
                Invoke-StaCall 'Export' 'Graph.SaveAsFormat (PNG)' {
                    if (-not $graph.SaveAsFormat($Stage, 6, $true)) { throw 'PNG export returned false.' }
                }
            }
            '.svg' { Invoke-StaCall 'Export' 'Graph.ExportSVG' { $graph.ExportSVG($Stage, 0, 0, 0) } }
        }
    } finally { Close-StaDocument $graph }
}
