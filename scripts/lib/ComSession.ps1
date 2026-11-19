function Write-StaEvent([string]$Event, [string]$Phase, [string]$Operation, $Details) {
    $entry = [ordered]@{
        time = [datetime]::UtcNow.ToString('o')
        event = $Event
        phase = $Phase
        operation = $Operation
        job = $staState.Job
        backend = $staState.Backend
        worker_pid = $PID
    }
    if ($Details) {
        foreach ($key in $Details.Keys) { $entry[$key] = $Details[$key] }
    }
    $line = ConvertTo-Json -InputObject $entry -Depth 12 -Compress
    [IO.File]::AppendAllText($staEventPath, $line + [Environment]::NewLine)
}

function New-StaError([string]$Phase, [string]$Operation, $Reason) {
    $hresult = $null
    if ($Reason -is [Management.Automation.ErrorRecord]) {
        $exception = $Reason.Exception.GetBaseException()
        if ($exception -is [Runtime.InteropServices.COMException] -or $exception -is [InvalidCastException]) {
            $hresult = '0x{0:X8}' -f $exception.HResult
        }
        $message = $exception.Message
    } else {
        $message = [string]$Reason
    }
    [ordered]@{
        phase = $Phase
        operation = $Operation
        backend = $staState.Backend
        hresult = $hresult
        message = $message
    }
}

function Invoke-StaCall {
    param([string]$Phase, [string]$Operation, [scriptblock]$Body, [switch]$Cleanup)

    $watch = [Diagnostics.Stopwatch]::StartNew()
    Write-StaEvent 'begin' $Phase $Operation
    try {
        & $Body
        Write-StaEvent 'end' $Phase $Operation @{ elapsed_ms = $watch.ElapsedMilliseconds }
    } catch {
        $failure = New-StaError $Phase $Operation $_
        $failure.elapsed_ms = $watch.ElapsedMilliseconds
        Write-StaEvent 'error' $Phase $Operation $failure
        if ($Cleanup) {
            $staState.CleanupFailures.Add($failure)
        } else {
            if (-not $staState.Failure) { $staState.Failure = $failure }
            throw
        }
    }
}

function Assert-StaCleanup {
    if ($staState.CleanupFailures.Count) {
        $failure = $staState.CleanupFailures[0]
        if (-not $staState.Failure) { $staState.Failure = $failure }
        throw "$($failure.operation) failed: $($failure.message)"
    }
}

function Release-StaCom($Object) {
    if ($null -ne $Object -and [Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Object) }
        catch {
            $failure = New-StaError 'Cleanup' 'ReleaseComObject' $_
            $staState.CleanupFailures.Add($failure)
            Write-StaEvent 'error' 'Cleanup' 'ReleaseComObject' $failure
        }
    }
}

function Close-StaDocument($Document, [switch]$LeaveForQuit) {
    if ($null -eq $Document) { return }
    if ($LeaveForQuit) {
        Invoke-StaCall 'Cleanup' 'Document.Saved' { $Document.Saved = $true } -Cleanup
    } else {
        Invoke-StaCall 'Cleanup' 'Document.Close(false)' { $Document.Close($false) } -Cleanup
    }
    Release-StaCom $Document
}

function Open-StaSession([string]$Backend) {
    $staState.Backend = $Backend
    $app = Invoke-StaCall 'Activate' 'New-Object' { ,(New-Object -ComObject "Statistica.$Backend") }
    $session = [pscustomobject]@{ Backend = $Backend; App = $app; Pid = $null }
    try {
        $serverPid = Invoke-StaCall 'Activate' 'Application.ProcessID' { [int]$app.ProcessID }
        if ($Backend -eq 'Application' -and $serverPid -in $staRequest.BaselinePids) {
            throw "COM returned a pre-existing STATISTICA process ($serverPid)."
        }
        $session.Pid = $serverPid
        $process = Get-Process -Id $serverPid -ErrorAction SilentlyContinue
        try {
            Write-StaEvent 'server' 'Activate' 'Application.ProcessID' @{
                pid = $serverPid
                session = if ($process) { $process.SessionId } else { $null }
                start_ticks = if ($process) { $process.StartTime.ToUniversalTime().Ticks } else { $null }
                shell = $PSVersionTable.PSVersion.ToString()
                apartment = [Threading.Thread]::CurrentThread.GetApartmentState().ToString()
                x64 = [Environment]::Is64BitProcess
            }
        } finally {
            if ($process) { $process.Dispose() }
        }
        $session
    } catch {
        # A pre-existing server can belong to another desktop session or user workflow.
        if ($Backend -eq 'Application' -and $session.Pid) {
            Invoke-StaCall 'Cleanup' 'Application.Quit' { $app.Quit() } -Cleanup
        }
        Release-StaCom $app
        throw
    }
}

function Clear-StaMacroDocuments($Session, [switch]$LeaveForQuit) {
    # A self-contained job owns all documents in its worker session.
    Invoke-StaCall 'Cleanup' 'Macro documents' {
        foreach ($name in 'Macros', 'Graphs', 'Spreadsheets', 'Workbooks', 'Reports') {
            $collection = $Session.App.$name
            try {
                for ($index = $collection.Count; $index -ge 1; $index--) {
                    $document = $collection.Item($index)
                    Close-StaDocument $document -LeaveForQuit:$LeaveForQuit
                }
            } finally {
                Release-StaCom $collection
            }
        }
    } -Cleanup
}

function Exit-StaSession {
    $session = $staState.Session
    if (-not $session) { return }
    $staState.Session = $null
    $staState.Backend = $session.Backend
    try {
        if ($session.Backend -eq 'Application') {
            Invoke-StaCall 'Cleanup' 'Application.Quit' { $session.App.Quit() } -Cleanup
        }
    } finally {
        Release-StaCom $session.App
    }
}

function Get-StaBackendCandidates($Job) {
    if ($staRequest.Backend -ne 'Auto') { return $staRequest.Backend }
    if ($Job.NeedsDesktop) { return 'Application' }
    if ($Job.Action -in 'Probe', 'Inspect', 'ExportCsv', 'RunMacro') { return 'Application', 'Library' }
    'Library', 'Application'
}

function Get-StaSession($Job, [string]$RequiredBackend) {
    $candidates = if ($RequiredBackend) { @($RequiredBackend) } else { @(Get-StaBackendCandidates $Job) }
    if ($staState.Session -and $staState.Session.Backend -in $candidates) {
        $staState.Backend = $staState.Session.Backend
        return $staState.Session
    }
    Exit-StaSession
    Assert-StaCleanup

    $errors = [Collections.Generic.List[string]]::new()
    foreach ($backend in $candidates) {
        if (-not $RequiredBackend -and $staState.ActivationFailures[$backend]) {
            $errors.Add("Statistica.$backend $($staState.ActivationFailures[$backend])")
            continue
        }
        $staState.Failure = $null
        try {
            $staState.Session = Open-StaSession $backend
            return $staState.Session
        } catch {
            $failure = $staState.Failure ?? (New-StaError 'Activate' 'Open-StaSession' $_)
            $reason = if ($failure.hresult) { "$($failure.hresult): $($failure.message)" } else { $failure.message }
            $staState.ActivationFailures[$backend] = $reason
            $errors.Add("Statistica.$backend $reason")
        }
    }
    $message = 'STATISTICA COM activation failed. ' + ($errors -join ' | ')
    if ($Job.Action -eq 'RunMacro' -and $Job.NeedsDesktop) {
        $message += ' This mode needs an active desktop spreadsheet. For an SVB that reads GetScriptArgument, use -PassInputAsScriptArgument.'
    }
    $failure = New-StaError 'Activate' 'Open-StaSession' $message
    if ($staState.Failure) { $failure.hresult = $staState.Failure.hresult }
    $staState.Failure = $failure
    throw $message
}

function Get-StaProbeResult($Session) {
    $version = Invoke-StaCall 'Probe' 'Application.Version' { [string]$Session.App.Version }
    [ordered]@{
        com = "Statistica.$($Session.Backend)"
        available = $true
        version = $version
        application_available = if ($Session.Backend -eq 'Application') { $true }
            elseif ($staState.ActivationFailures['Application']) { $false } else { $null }
        application_error = $staState.ActivationFailures['Application']
    }
}

function Get-StaRegistration([string]$Backend, [Microsoft.Win32.RegistryView]$View) {
    $root = [Microsoft.Win32.RegistryKey]::OpenBaseKey('ClassesRoot', $View)
    $class = $server = $null
    try {
        $class = $root.OpenSubKey("Statistica.$Backend\CLSID")
        if (-not $class) { return }
        $id = $class.GetValue('')
        $kind = if ($Backend -eq 'Application') { 'LocalServer32' } else { 'InprocServer32' }
        $server = $root.OpenSubKey("CLSID\$id\$kind")
        [ordered]@{
            view = $View.ToString()
            clsid = $id
            kind = $kind
            server = if ($server) { $server.GetValue('') } else { $null }
        }
    } finally {
        foreach ($key in $server, $class, $root) { if ($key) { $key.Dispose() } }
    }
}

function Invoke-StaDetailedProbe($Job) {
    Exit-StaSession
    Assert-StaCleanup
    $details = [ordered]@{}
    $result = $null
    foreach ($backend in @(Get-StaBackendCandidates $Job)) {
        $component = [ordered]@{
            available = $false; version = $null; error = $null; pid = $null
            registration = @(Get-StaRegistration $backend 'Registry64'; Get-StaRegistration $backend 'Registry32')
        }
        $staState.Failure = $null
        if (-not $staState.ActivationFailures[$backend]) {
            try {
                $staState.Session = Open-StaSession $backend
                $probe = Get-StaProbeResult $staState.Session
                $component.available = $true
                $component.version = $probe.version
                $component.pid = $staState.Session.Pid
                if (-not $result) { $result = $probe }
            } catch {
                $failure = $staState.Failure ?? (New-StaError 'Probe' 'Open-StaSession' $_)
                $staState.ActivationFailures[$backend] = ('{0} {1}' -f $failure.hresult, $failure.message).Trim()
            } finally {
                Exit-StaSession
            }
        }
        $component.error = $staState.ActivationFailures[$backend]
        $details[$backend] = $component
        Assert-StaCleanup
    }
    $staState.Failure = $null
    if (-not $result) {
        $result = [ordered]@{ com = $null; available = $false; version = $null
            application_available = if ($details.Contains('Application')) { $false } else { $null }
            application_error = $staState.ActivationFailures['Application'] }
    }
    $result['components'] = $details
    $result['host'] = [ordered]@{
        pwsh = $PSVersionTable.PSVersion.ToString()
        x64 = [Environment]::Is64BitProcess
        apartment = [Threading.Thread]::CurrentThread.GetApartmentState().ToString()
        session = [Diagnostics.Process]::GetCurrentProcess().SessionId
    }
    $result
}
