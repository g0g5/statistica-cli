$staCommandOptions = @{
    Probe = @('Detailed')
    Inspect = @('InputPath')
    ExportCsv = @('InputPath', 'OutputPath', 'Columns', 'FirstRow', 'LastRow', 'DisplayText', 'Force')
    RunMacro = @('InputPath', 'OutputPath', 'MacroPath', 'PassInputAsScriptArgument', 'VerifyCells', 'Tolerance', 'Force')
    EditSta = @('InputPath', 'OutputPath', 'SpecPath', 'Force')
    CreateSta = @('OutputPath', 'SpecPath', 'Force')
    ImportCsv = @('InputPath', 'OutputPath', 'SpecPath', 'MissingToken', 'Force')
    CreateGraph = @('InputPath', 'OutputPath', 'GraphType', 'XColumn', 'YColumns', 'ConnectPoints', 'Force')
    ExportGraph = @('InputPath', 'OutputPath', 'Force')
    Screenshot = @('InputPath', 'OutputPath', 'CropToDocument', 'CaptureTimeoutSeconds', 'Force')
    ListWorkbook = @('InputPath')
    ExtractWorkbookItem = @('InputPath', 'OutputPath', 'ItemIndex', 'ItemPath', 'Force')
    BuildWorkbook = @('OutputPath', 'DocumentPaths', 'Force')
}

function Resolve-StaFilePath {
    param([string]$Value, [string[]]$Extensions, [string]$Option, [string]$BasePath,
        [switch]$Existing, [hashtable]$PriorOutputs)

    if ([string]::IsNullOrWhiteSpace($Value)) { throw "Supply -$Option." }
    $provider = $null
    $drive = $null
    $path = $Value
    if ($BasePath -and -not [IO.Path]::IsPathRooted($path) -and $path -notmatch '^[^:]+:') {
        $path = Join-Path $BasePath $path
    }
    $path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($path, [ref]$provider, [ref]$drive)
    if ($provider.Name -ne 'FileSystem' -or [IO.Path]::GetExtension($path) -notin $Extensions) {
        throw "-$Option must be a file with extension $($Extensions -join ', ')."
    }
    if ($Existing -and -not ($PriorOutputs -and $PriorOutputs.ContainsKey($path)) -and -not [IO.File]::Exists($path)) {
        throw ('-{0} file does not exist: {1}' -f $Option, $path)
    }
    $path
}

function Test-StaInteger($Value, [int]$Minimum, [int]$Maximum = [int]::MaxValue) {
    ($Value -is [int] -or $Value -is [long]) -and $Value -ge $Minimum -and $Value -le $Maximum
}

function New-StaJob {
    param([hashtable]$Entry, [int]$Id, [string]$BasePath, [hashtable]$PriorOutputs,
        [string]$Backend, [bool]$FromManifest)

    $command = [string]$Entry['Action']
    if (-not $staCommandOptions.ContainsKey($command)) { throw "Unsupported action '$command' in job $Id." }
    $command = $staCommandOptions.Keys | Where-Object { $_ -eq $command } | Select-Object -First 1
    $options = @{}
    foreach ($option in $staCommandOptions[$command]) {
        if ($Entry.ContainsKey($option)) { $options[$option] = $Entry[$option] }
    }
    if ($FromManifest) {
        foreach ($option in $Entry.Keys) {
            if ($option -ne 'Action' -and $option -notin $staCommandOptions[$command]) {
                throw "Unknown option '$option' for job $Id ($command)."
            }
        }
    }
    foreach ($option in 'DisplayText', 'PassInputAsScriptArgument', 'Detailed', 'Force', 'ConnectPoints', 'CropToDocument') {
        if ($options.ContainsKey($option)) {
            if ($options[$option] -isnot [bool] -and $options[$option] -isnot [switch]) {
                throw "-$option must be Boolean."
            }
            $options[$option] = [bool]$options[$option]
        }
    }

    if ($command -eq 'ExportCsv') {
        $options['FirstRow'] = $options['FirstRow'] ?? 1
        $options['LastRow'] = $options['LastRow'] ?? 0
        if (-not (Test-StaInteger $options.FirstRow 1) -or -not (Test-StaInteger $options.LastRow 0)) {
            throw 'FirstRow and LastRow must be in the supported integer range.'
        }
        if ($options.ContainsKey('Columns') -and $options.Columns -isnot [string]) {
            throw 'Columns must be a comma-separated string.'
        }
    }
    if ($command -eq 'Screenshot') {
        $options['CaptureTimeoutSeconds'] = $options['CaptureTimeoutSeconds'] ?? 8
        if (-not (Test-StaInteger $options.CaptureTimeoutSeconds 1 120)) {
            throw 'CaptureTimeoutSeconds must be an integer from 1 to 120.'
        }
    }

    $needsDesktop = $command -eq 'Screenshot' -or
        ($command -eq 'RunMacro' -and -not $options['PassInputAsScriptArgument'])
    if ($needsDesktop -and $Backend -eq 'Library') {
        throw "$command requires -Backend Application or Auto in this mode."
    }

    $inputs = [Collections.Generic.List[string]]::new()
    if ('InputPath' -in $staCommandOptions[$command]) {
        $extensions = switch ($command) {
            ImportCsv { '.csv' }
            ExportGraph { '.stg' }
            Screenshot { '.sta'; '.stg'; '.stw' }
            ListWorkbook { '.stw' }
            ExtractWorkbookItem { '.stw' }
            default { '.sta' }
        }
        $options.InputPath = Resolve-StaFilePath $options['InputPath'] $extensions InputPath $BasePath -Existing -PriorOutputs $PriorOutputs
        $inputs.Add($options.InputPath)
    }
    if ('MacroPath' -in $staCommandOptions[$command]) {
        $options.MacroPath = Resolve-StaFilePath $options['MacroPath'] '.svb' MacroPath $BasePath -Existing
        $inputs.Add($options.MacroPath)
    }
    if ('SpecPath' -in $staCommandOptions[$command]) {
        $options.SpecPath = Resolve-StaFilePath $options['SpecPath'] '.json' SpecPath $BasePath -Existing
        $inputs.Add($options.SpecPath)
    }
    if ($command -eq 'BuildWorkbook') {
        if (@($options['DocumentPaths'] | Where-Object { $_ }).Count -eq 0) {
            throw 'Supply at least one -DocumentPaths entry.'
        }
        $options.DocumentPaths = @(
            foreach ($path in $options.DocumentPaths) {
                $resolved = Resolve-StaFilePath $path @('.sta', '.stg') DocumentPaths $BasePath -Existing -PriorOutputs $PriorOutputs
                $inputs.Add($resolved)
                $resolved
            }
        )
    }
    if ($command -eq 'ExtractWorkbookItem') {
        $byIndex = $options.ContainsKey('ItemIndex')
        $byPath = -not [string]::IsNullOrWhiteSpace($options['ItemPath'])
        if ($byIndex -eq $byPath -or ($byIndex -and -not (Test-StaInteger $options.ItemIndex 1))) {
            throw 'Supply exactly one positive -ItemIndex or one -ItemPath.'
        }
    }
    if ($command -eq 'CreateGraph') {
        if ($options['GraphType'] -notin 'Line', 'Scatter') { throw 'Supply -GraphType Line or Scatter.' }
        if (-not $options['YColumns']) { throw 'Supply at least one -YColumns column number.' }
        foreach ($column in @($options.YColumns)) {
            if (-not (Test-StaInteger $column 1)) { throw 'YColumns must contain positive integer column numbers.' }
        }
        if (@($options.YColumns | Select-Object -Unique).Count -ne @($options.YColumns).Count) {
            throw 'YColumns cannot include duplicate columns.'
        }
        if ($options.GraphType -eq 'Line') {
            if ($options.ContainsKey('XColumn') -or $options['ConnectPoints']) {
                throw 'XColumn and ConnectPoints apply to scatterplots.'
            }
        } elseif (-not (Test-StaInteger $options['XColumn'] 1) -or $options.XColumn -in $options.YColumns) {
            throw 'Supply a positive -XColumn distinct from -YColumns for a scatterplot.'
        }
    }
    if ($options.ContainsKey('VerifyCells') -and $options.VerifyCells -isnot [System.Collections.IDictionary]) {
        throw 'VerifyCells must map row,column keys to values.'
    }
    if ($options.ContainsKey('Tolerance') -and
        ($options.Tolerance -isnot [ValueType] -or $options.Tolerance -is [bool] -or
         -not [double]::IsFinite($options.Tolerance) -or $options.Tolerance -lt 0)) {
        throw 'Tolerance must be a finite nonnegative number.'
    }
    if ($options.ContainsKey('MissingToken') -and $options.MissingToken -isnot [string]) {
        throw 'MissingToken must be a string.'
    }

    $destination = $null
    $stage = $null
    if ('OutputPath' -in $staCommandOptions[$command]) {
        $extensions = switch ($command) {
            ExportCsv { '.csv' }
            CreateGraph { '.stg' }
            ExportGraph { '.png'; '.svg' }
            Screenshot { '.png' }
            BuildWorkbook { '.stw' }
            ExtractWorkbookItem { '.sta'; '.stg' }
            default { '.sta' }
        }
        $destination = Resolve-StaFilePath $options['OutputPath'] $extensions OutputPath $BasePath
        $options.Remove('OutputPath')
        $directory = [IO.Path]::GetDirectoryName($destination)
        if (-not [IO.Directory]::Exists($directory)) { throw "Output directory does not exist: $directory" }
        if ($destination -in $inputs -or $PriorOutputs.ContainsKey($destination)) {
            throw "Output must be distinct from its inputs and earlier job outputs: $destination"
        }
        if ([IO.Directory]::Exists($destination) -or ([IO.File]::Exists($destination) -and -not $options['Force'])) {
            throw "Output already exists: $destination"
        }
        $stage = Join-Path $directory ('.statistica-' + [guid]::NewGuid().ToString('N') + [IO.Path]::GetExtension($destination))
    }

    $dependencies = @(
        foreach ($path in $inputs) {
            if ($PriorOutputs.ContainsKey($path)) { $PriorOutputs[$path].Id }
        }
    ) | Sort-Object -Unique
    $job = [ordered]@{
        Id = $Id
        Action = $command
        Options = $options
        Stage = $stage
        Destination = $destination
        Dependencies = @($dependencies | Where-Object { $null -ne $_ })
        Inputs = $inputs.ToArray()
        NeedsDesktop = $needsDesktop
    }
    if ($destination) { $PriorOutputs[$destination] = $job }
    $job
}

function New-StaRequest($Settings, [hashtable]$Bound) {
    $jobs = [Collections.Generic.List[object]]::new()
    $prior = @{}
    $batch = $Settings.Action -eq 'Batch'
    $manifest = $null
    if ($batch) {
        $manifest = Resolve-StaFilePath $Bound['ManifestPath'] '.json' ManifestPath -Existing
        $data = Get-Content -LiteralPath $manifest -Raw | ConvertFrom-Json -AsHashtable
        if ($data -isnot [System.Collections.IDictionary] -or $data['jobs'] -isnot [array] -or $data.jobs.Count -eq 0) {
            throw 'Manifest must be a JSON object with a nonempty jobs array.'
        }
        $basePath = [IO.Path]::GetDirectoryName($manifest)
        foreach ($entry in $data.jobs) {
            if ($entry -isnot [System.Collections.IDictionary]) { throw 'Each batch job must be an object.' }
            $normalized = @{}
            foreach ($key in $entry.Keys) { $normalized[$key] = $entry[$key] }
            $jobs.Add((New-StaJob $normalized ($jobs.Count + 1) $basePath $prior $Settings.Backend $true))
        }
    } else {
        if ($Bound['PassInputAsScriptArgument'] -and $Settings.Action -ne 'RunMacro') {
            throw '-PassInputAsScriptArgument is only valid with RunMacro.'
        }
        if ($Bound.ContainsKey('ManifestPath') -or $Bound['ContinueOnError']) {
            throw 'ManifestPath and ContinueOnError apply to Batch.'
        }
        $jobs.Add((New-StaJob $Bound 1 '' $prior $Settings.Backend $false))
    }

    $diagnostic = $null
    if ($Settings.DiagnosticPath) {
        $diagnostic = Resolve-StaFilePath $Settings.DiagnosticPath '.jsonl' DiagnosticPath
        if (-not [IO.Directory]::Exists([IO.Path]::GetDirectoryName($diagnostic)) -or
            [IO.File]::Exists($diagnostic) -or [IO.Directory]::Exists($diagnostic) -or
            $diagnostic -eq $manifest -or
            ($jobs | Where-Object { $diagnostic -in $_.Inputs -or $diagnostic -eq $_.Destination })) {
            throw "DiagnosticPath must identify a new file in an existing directory: $diagnostic"
        }
    }
    [ordered]@{
        Backend = $Settings.Backend
        TimeoutSeconds = $Settings.TimeoutSeconds
        DiagnosticPath = $diagnostic
        Verbose = $Settings.Verbose
        Batch = $batch
        ContinueOnError = $Settings.ContinueOnError
        Jobs = $jobs.ToArray()
        ActivationFailures = @{}
    }
}
