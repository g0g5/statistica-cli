#Requires -Version 7.5
param([Parameter(Mandatory)] [string]$RequestPath)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

$staDirectory = [IO.Path]::GetDirectoryName($RequestPath)
$staEventPath = Join-Path $staDirectory 'events.jsonl'
$staRequest = Get-Content -LiteralPath $RequestPath -Raw | ConvertFrom-Json -AsHashtable -DateKind String
$staState = @{
    Job = 0; Backend = $null; Session = $null; Failure = $null
    CleanupFailures = [Collections.Generic.List[object]]::new()
    ActivationFailures = $staRequest.ActivationFailures
}
foreach ($part in 'ComSession', 'OutputFiles', 'Spreadsheets', 'Macros', 'Graphs', 'Workbooks', 'WindowCapture') {
    . (Join-Path $PSScriptRoot "lib/$part.ps1")
}

$jobResults = @{}
foreach ($job in $staRequest.Jobs) {
    $staState.Job = $job.Id
    $staState.Failure = $null
    $staState.CleanupFailures.Clear()
    [IO.File]::WriteAllText((Join-Path $staDirectory "job-$($job.Id).started"), '')
    Write-StaEvent 'job_begin' 'Job' $job.Action

    $result = $null
    $backend = $null
    $blocked = @($job.Dependencies | Where-Object { $jobResults[[int]$_].error })
    try {
        if ($blocked.Count) {
            $staState.Failure = New-StaError 'Dependency' $job.Action "Dependency jobs failed: $($blocked -join ', ')."
            throw $staState.Failure.message
        }

        $options = @{} + $job.Options
        foreach ($dependency in $job.Dependencies) {
            $prior = $staRequest.Jobs[$dependency - 1]
            if ($options['InputPath'] -eq $prior.Destination) { $options.InputPath = $prior.Stage }
            if ($options['DocumentPaths']) {
                $options.DocumentPaths = @(
                    foreach ($path in $options.DocumentPaths) {
                        if ($path -eq $prior.Destination) { $prior.Stage } else { $path }
                    }
                )
            }
        }

        if ($job.Action -eq 'Probe' -and $options['Detailed']) {
            $result = Invoke-StaDetailedProbe $job
            $backend = $result.com -replace '^Statistica\.'
        } else {
            # A visible job owns its entire desktop instance until Quit.
            if ($job.NeedsDesktop) {
                if ($staState.Session -and $staState.Session.Backend -eq 'Application') { Exit-StaSession }
                Assert-StaCleanup
            }
            $session = Get-StaSession $job
            $backend = $session.Backend
            $result = switch ($job.Action) {
                Probe { Get-StaProbeResult $session }
                Inspect { Get-StaSpreadsheetInfo $session $options.InputPath $job.Options.InputPath }
                ExportCsv { Export-StaCsv $session $options $job.Stage }
                RunMacro { Invoke-StaMacro $session $options $job }
                EditSta { Edit-StaSpreadsheet $session $options $job.Stage }
                CreateSta { New-StaSpreadsheet $session $options $job.Stage }
                ImportCsv { New-StaSpreadsheet $session $options $job.Stage -Import }
                CreateGraph { New-StaGraph $session $options $job.Stage }
                ExportGraph { Export-StaGraph $session $options $job.Stage }
                Screenshot { Save-StaWindowCapture $session $options $job.Stage }
                ListWorkbook { Get-StaWorkbookInfo $session $options.InputPath $job.Options.InputPath }
                ExtractWorkbookItem { Export-StaWorkbookItem $session $options $job.Stage }
                BuildWorkbook { New-StaWorkbook $session $options.DocumentPaths $job.Stage $job.Options.DocumentPaths }
            }
        }
        if ($job.Stage) { Confirm-StaFile $job.Stage }
    } catch {
        if (-not $staState.Failure) {
            $staState.Failure = New-StaError 'Job' $job.Action $_
            Write-StaEvent 'error' 'Job' $job.Action $staState.Failure
        }
    } finally {
        if ($job.NeedsDesktop -and $staState.Session -and $staState.Session.Backend -eq 'Application') {
            Exit-StaSession
        }
    }

    if (-not $staState.Failure -and $staState.CleanupFailures.Count) {
        $staState.Failure = $staState.CleanupFailures[0]
    }
    $record = [ordered]@{
        backend = $backend ?? $staState.Backend
        result = $result
        error = $staState.Failure
    }
    $jobResults[[int]$job.Id] = $record
    $checkpoint = Join-Path $staDirectory "job-$($job.Id).json"
    $json = ConvertTo-Json -InputObject $record -Depth 24 -Compress -WarningAction Stop
    [IO.File]::WriteAllText("$checkpoint.tmp", $json)
    [IO.File]::Move("$checkpoint.tmp", $checkpoint)
    Write-StaEvent 'job_end' 'Job' $job.Action @{ succeeded = -not [bool]$record.error }

    if ($staState.CleanupFailures.Count -or ($record.error -and -not $staRequest.ContinueOnError)) { break }
}
$staState.Failure = $null
$staState.CleanupFailures.Clear()
Exit-StaSession
Assert-StaCleanup
