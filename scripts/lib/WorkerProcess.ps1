function Get-StaWorkerEvents([string]$Path) {
    if (-not [IO.File]::Exists($Path)) { return }
    foreach ($line in Get-Content -LiteralPath $Path -ErrorAction Stop) {
        try { ConvertFrom-Json -InputObject $line -AsHashtable -DateKind String }
        catch { Write-Debug 'Ignoring an incomplete diagnostic record.' }
    }
}

function Get-StaDesktopSnapshot {
    try {
        Get-CimInstance Win32_Process -Filter "Name = 'statist.exe'" -OperationTimeoutSec 3 |
            ForEach-Object {
                [ordered]@{
                    pid = [int]$_.ProcessId
                    session = [int]$_.SessionId
                    start = $_.CreationDate.ToUniversalTime().ToString('o')
                    command_line = $_.CommandLine
                }
            }
    } catch {
        Write-Verbose "Unable to read STATISTICA process details: $($_.Exception.Message)"
    }
}

function Get-StaLastCall($Events, [bool]$TimedOut) {
    if (-not $TimedOut) { return $Events | Where-Object event -eq 'error' | Select-Object -Last 1 }
    $open = [Collections.Generic.Stack[object]]::new()
    foreach ($event in $Events) {
        if ($event.event -eq 'begin') {
            $open.Push($event)
        } elseif ($event.event -in 'end', 'error' -and $open.Count) {
            $call = $open.Peek()
            if ($call.phase -eq $event.phase -and $call.operation -eq $event.operation) { [void]$open.Pop() }
        }
    }
    if ($open.Count) { $open.Peek() }
    else { $Events | Where-Object event -eq 'job_begin' | Select-Object -Last 1 }
}

function Stop-StaOwnedServers($Events, [int[]]$Baseline, [datetime]$Started) {
    $servers = @($Events | Where-Object { $_.event -eq 'server' -and $_.backend -eq 'Application' })
    if (-not $servers.Count) { return }

    $candidates = @(Get-StaDesktopSnapshot)
    foreach ($server in $servers) {
        if ($server.pid -in $Baseline -or -not $server['start_ticks']) { continue }
        $candidate = $candidates | Where-Object {
            $_.pid -eq $server.pid -and $_.session -eq $server.session -and
            $_.command_line -match '(?i)statist\.exe"?\s+-Embedding\s*$'
        } | Select-Object -First 1
        if (-not $candidate) { continue }

        $process = Get-Process -Id $server.pid -ErrorAction SilentlyContinue
        if ($process) {
            try {
                if ($process.ProcessName -eq 'statist' -and $process.StartTime.ToUniversalTime() -ge $Started -and
                    $process.StartTime.ToUniversalTime().Ticks -eq $server.start_ticks) {
                    $process.Kill()
                    if (-not $process.WaitForExit(3000)) { throw 'The process did not exit after termination.' }
                    Write-Verbose "Ended STATISTICA process $($server.pid) owned by the failed worker."
                    $server.pid
                }
            } catch {
                Write-Warning "Could not end the owned STATISTICA process $($server.pid): $($_.Exception.Message)"
            } finally {
                $process.Dispose()
            }
        }
    }
}

function Invoke-StaWorker($Request, [string]$WorkerPath, [string]$Directory) {
    [void][IO.Directory]::CreateDirectory($Directory)
    $requestPath = Join-Path $Directory 'request.json'
    $eventsPath = Join-Path $Directory 'events.jsonl'
    $baseline = @(Get-Process -Name statist -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $Request['BaselinePids'] = $baseline
    $Request | ConvertTo-Json -Depth 30 -WarningAction Stop | Set-Content -LiteralPath $requestPath -Encoding utf8

    $info = [Diagnostics.ProcessStartInfo]::new()
    $info.FileName = Join-Path $PSHOME 'pwsh.exe'
    $info.UseShellExecute = $false
    $info.CreateNoWindow = $true
    $info.RedirectStandardOutput = $true
    $info.RedirectStandardError = $true
    foreach ($argument in '-NoLogo', '-NoProfile', '-NonInteractive', '-STA', '-File', $WorkerPath, $requestPath) {
        $info.ArgumentList.Add($argument)
    }

    $started = [datetime]::UtcNow
    $child = [Diagnostics.Process]::Start($info)
    $watch = [Diagnostics.Stopwatch]::StartNew()
    $timeout = $false
    $lastJob = 0
    $events = @()
    $terminated = @()
    try {
        $stdout = $child.StandardOutput.ReadToEndAsync()
        $stderr = $child.StandardError.ReadToEndAsync()
        while (-not $child.WaitForExit(125)) {
            while ([IO.File]::Exists((Join-Path $Directory "job-$($lastJob + 1).started"))) {
                $lastJob++
                $watch.Restart()
            }
            if ($watch.Elapsed.TotalSeconds -ge $Request.TimeoutSeconds) {
                $timeout = $true
                break
            }
        }
    } finally {
        if (-not $child.HasExited) {
            try {
                $child.Kill()
                [void]$child.WaitForExit(5000)
            } catch [InvalidOperationException] {
                if (-not $child.HasExited) { throw }
            }
            $events = @(Get-StaWorkerEvents $eventsPath)
            $terminated = @(Stop-StaOwnedServers $events $baseline $started)
        }
    }

    try {
        $events = @(Get-StaWorkerEvents $eventsPath)
        if (-not $timeout -and $child.ExitCode -ne 0) {
            $terminated = @(Stop-StaOwnedServers $events $baseline $started)
        }
        $records = @{}
        foreach ($job in $Request.Jobs) {
            $path = Join-Path $Directory "job-$($job.Id).json"
            if ([IO.File]::Exists($path)) {
                $records[[int]$job.Id] = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json -AsHashtable -DateKind String
            }
        }

        $failure = $null
        if ($timeout -or $child.ExitCode -ne 0) {
            $lastCall = Get-StaLastCall $events $timeout
            $failure = [ordered]@{
                status = if ($timeout) { 'timed_out' } else { 'failed' }
                phase = if ($lastCall) { $lastCall.phase } else { 'Worker' }
                operation = if ($lastCall) { $lastCall.operation } else { 'Start-Process' }
                backend = if ($lastCall) { $lastCall.backend } else { $null }
                hresult = if ($lastCall) { $lastCall['hresult'] } else { $null }
                job = if ($lastCall) { $lastCall.job } else { 0 }
                message = if ($timeout) { "STATISTICA worker exceeded $($Request.TimeoutSeconds) seconds." }
                    elseif ($lastCall -and $lastCall['message']) { $lastCall.message }
                    else { "STATISTICA worker exited with code $($child.ExitCode)." }
            }
            if (-not $timeout -and $stderr.IsCompletedSuccessfully -and $stderr.Result -and
                (-not $lastCall -or -not $lastCall['message'])) {
                $failure.message += ' ' + $stderr.Result.Trim()
            }
        }
        $remaining = @()
        if ($failure -or ($events | Where-Object { $_.phase -eq 'Activate' -and $_.event -eq 'error' })) {
            $remaining = @(Get-StaDesktopSnapshot)
        }
        foreach ($server in $events | Where-Object { $_.event -eq 'server' -and $_.backend -eq 'Application' }) {
            if ($server.pid -notin $baseline -and -not $failure -and
                (Get-Process -Id $server.pid -ErrorAction SilentlyContinue | Where-Object {
                    $_.ProcessName -eq 'statist' -and $server['start_ticks'] -and
                    $_.StartTime.ToUniversalTime().Ticks -eq $server.start_ticks
                })) {
                Write-Verbose "STATISTICA process $($server.pid) was still present after worker exit."
            }
        }
        [pscustomobject]@{
            Events = $events
            Records = $records
            Failure = $failure
            Remaining = $remaining
            EndedServerPids = $terminated
            WorkerId = $child.Id
            WorkerStdout = if ($stdout.IsCompletedSuccessfully -and $stdout.Result) { $stdout.Result.Trim() } else { $null }
        }
    } finally {
        $child.Dispose()
    }
}

function Format-StaFailure($Failure) {
    $code = ''
    if ($Failure['hresult']) { $code = ' ' + $Failure.hresult }
    '[{0}/{1}{2}] {3}' -f $Failure.phase, $Failure.operation, $code, $Failure.message
}

function Write-StaRunDiagnostics($Request, $Run) {
    foreach ($event in $Run.Events) {
        if ($Request.Verbose -and $event.event -in 'begin', 'end', 'error', 'server', 'job_begin', 'job_end') {
            $detail = if ($event['message']) { ' ' + $event.message } else { '' }
            $duration = if ($event['elapsed_ms'] -ne $null) { ' ({0} ms)' -f $event.elapsed_ms } else { '' }
            Write-Verbose ("Job {0} {1} {2}/{3} {4}{5}{6}" -f
                $event.job, $event.event, $event.phase, $event.operation, $event.backend, $duration, $detail)
        }
    }
    if ($Request.DiagnosticPath) {
        $events = @($Run.Events)
        $events += [ordered]@{
            event = 'worker_exit'
            time = [datetime]::UtcNow.ToString('o')
            worker_pid = $Run.WorkerId
            failure = $Run.Failure
            desktop_processes = $Run.Remaining
            ended_server_pids = $Run.EndedServerPids
            worker_stdout = $Run.WorkerStdout
        }
        $lines = foreach ($event in $events) { ConvertTo-Json -InputObject $event -Depth 12 -Compress }
        [IO.File]::AppendAllLines($Request.DiagnosticPath, [string[]]$lines)
    }
}

function Write-StaPublishDiagnostic($Request, $Job, $Entry) {
    if (-not $Request.DiagnosticPath -or -not $Job.Destination) { return }
    $event = [ordered]@{
        time = [datetime]::UtcNow.ToString('o'); event = 'publish'; phase = 'Publish'
        operation = 'File.Move'; job = $Job.Id; status = $Entry.status
        destination = $Job.Destination; message = $Entry.error
    }
    [IO.File]::AppendAllText($Request.DiagnosticPath,
        (ConvertTo-Json -InputObject $event -Compress) + [Environment]::NewLine)
}

function Invoke-StaRequest($Request, [string]$WorkerPath) {
    if ($Request.DiagnosticPath) {
        $log = [IO.File]::Open($Request.DiagnosticPath, 'CreateNew', 'Write', 'Read')
        $log.Dispose()
    }
    $directory = Join-Path ([IO.Path]::GetTempPath()) ('statistica-cli-' + [guid]::NewGuid().ToString('N'))
    [void][IO.Directory]::CreateDirectory($directory)
    try {
        $run = Invoke-StaWorker $Request $WorkerPath (Join-Path $directory 'initial')
        Write-StaRunDiagnostics $Request $run

        # A call that never returned cannot take the worker's normal activation fallback.
        $failure = $run.Failure
        $firstJob = $Request.Jobs[0]
        if ($failure -and $failure.status -eq 'timed_out' -and $failure.phase -eq 'Activate' -and
            $failure.operation -in 'New-Object', 'Application.ProcessID' -and $failure.job -eq 1 -and $run.Records.Count -eq 0 -and
            $Request.Backend -eq 'Auto' -and -not $firstJob.NeedsDesktop) {
            $other = if ($failure.backend -eq 'Application') { 'Library' } else { 'Application' }
            $otherAttempted = $run.Events | Where-Object { $_.phase -eq 'Activate' -and $_.operation -eq 'New-Object' -and $_.backend -eq $other }
            if ($firstJob.Options['Detailed'] -or -not $otherAttempted) {
                $Request.ActivationFailures[$failure.backend] = Format-StaFailure $failure
                Write-Verbose "Activation timed out for Statistica.$($failure.backend). Trying Statistica.$other."
                $run = Invoke-StaWorker $Request $WorkerPath (Join-Path $directory 'fallback')
                Write-StaRunDiagnostics $Request $run
            }
        }

        if (-not $Request.Batch -and $run.Failure) { throw (Format-StaFailure $run.Failure) }
        $results = [Collections.Generic.List[object]]::new()
        $byId = @{}
        $stopPublishing = $false
        $workerFailureAssigned = $false
        foreach ($job in $Request.Jobs) {
            $entry = [ordered]@{
                id = $job.Id; action = $job.Action; status = 'skipped'
                backend = $null; output = $null; result = $null; error = $null
            }
            $record = $run.Records[[int]$job.Id]
            $blocked = @($job.Dependencies | Where-Object { $byId[[int]$_].status -ne 'succeeded' })
            if ($blocked.Count) {
                $entry.error = "Dependency jobs did not publish successfully: $($blocked -join ', ')."
            } elseif ($stopPublishing) {
                $entry.error = 'A previous job failed.'
            } elseif ($record) {
                $entry.backend = $record.backend
                if ($record.error) {
                    $entry.status = 'failed'
                    $entry.error = Format-StaFailure $record.error
                } else {
                    try {
                        if ($job.Destination) {
                            if (-not [IO.File]::Exists($job.Stage)) { throw 'Verified temporary output is missing.' }
                            [IO.File]::Move($job.Stage, $job.Destination, [bool]$job.Options['Force'])
                            $entry.output = $job.Destination
                        }
                        $entry.result = $record.result
                        $entry.status = 'succeeded'
                    } catch {
                        $entry.status = 'failed'
                        $entry.error = '[Publish/File.Move] ' + $_.Exception.GetBaseException().Message
                    }
                }
            } elseif ($run.Failure -and -not $workerFailureAssigned) {
                $workerFailureAssigned = $true
                $entry.status = $run.Failure.status
                $entry.backend = $run.Failure.backend
                $entry.error = Format-StaFailure $run.Failure
            } else {
                $entry.error = 'The worker did not complete this job.'
            }
            if (-not $Request.ContinueOnError -and $entry.status -ne 'succeeded') { $stopPublishing = $true }
            Write-StaPublishDiagnostic $Request $job $entry
            $byId[[int]$job.Id] = $entry
            $results.Add($entry)
        }

        if ($Request.Batch) {
            $failed = @($results | Where-Object { $_.status -ne 'succeeded' }).Count
            [ordered]@{ jobs = $results.ToArray(); worker_error = $run.Failure } | ConvertTo-Json -Depth 24
            if ($failed -or $run.Failure) { throw "STATISTICA batch did not finish successfully ($failed incomplete jobs)." }
        } else {
            $entry = $results[0]
            if ($entry.error) { throw $entry.error }
            if ($entry.output) { $entry.output }
            else { ConvertTo-Json -InputObject $entry.result -Depth 24 }
        }
    } finally {
        foreach ($job in $Request.Jobs) {
            if ($job.Stage -and [IO.File]::Exists($job.Stage)) {
                try { [IO.File]::Delete($job.Stage) }
                catch { Write-Warning "Could not remove temporary output $($job.Stage): $($_.Exception.Message)" }
            }
        }
        # Worker files have fixed names and are written only inside the unique run directory.
        foreach ($name in 'initial', 'fallback') {
            $attempt = Join-Path $directory $name
            if ([IO.Directory]::Exists($attempt)) {
                try {
                    foreach ($file in [IO.Directory]::EnumerateFiles($attempt)) { [IO.File]::Delete($file) }
                    [IO.Directory]::Delete($attempt)
                } catch { Write-Warning "Could not remove worker records in $attempt." }
            }
        }
        try { [IO.Directory]::Delete($directory) }
        catch { Write-Warning "Could not remove the worker directory $directory." }
    }
}
