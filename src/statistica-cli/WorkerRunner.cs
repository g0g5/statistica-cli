using System.ComponentModel;
using System.Diagnostics;
using System.IO.Pipes;
using System.Text;
using System.Text.Json;

namespace StatisticaCli;

internal sealed record WorkerRun(WorkerResult? Result, WorkerFailure? Failure, WorkerEvent[] Events, WorkerExit Exit, JobRecord[] Records);

internal static class WorkerRunner
{
    internal static async Task<WorkerRun> RunWithFallbackAsync(string executable, WorkerRequest request,
        int timeoutSeconds, TextWriter? diagnostics, bool verbose)
    {
        var run = await RunAsync(executable, request, timeoutSeconds, diagnostics, verbose);
        if (request.Backend == Backend.Auto && request.Jobs?.FirstOrDefault()?.NeedsDesktop != true &&
            run.Records.Length == 0 && run.Failure is { Status: "timed_out", Phase: "Activate",
                Operation: "CoCreateInstance" or "Application.ProcessID", Backend: "Application" or "Library", Job: 1 } failure)
        {
            var other = failure.Backend == "Application" ? "Library" : "Application";
            bool otherAttempted = run.Events.Any(entry =>
                entry is { Event: "begin", Phase: "Activate", Operation: "CoCreateInstance" } && entry.Backend == other);
            if (request.Detailed || !otherAttempted)
            {
                request = request with
                {
                    BaselinePids = StaProcesses.Baseline(),
                    ApplicationError = failure.Backend == "Application" ? failure.Format() : null,
                    LibraryError = failure.Backend == "Library" ? failure.Format() : null
                };
                run = await RunAsync(executable, request, timeoutSeconds, diagnostics, verbose);
            }
        }
        return run;
    }

    internal static async Task<WorkerRun> RunAsync(string executable, WorkerRequest request, int timeoutSeconds,
        TextWriter? diagnostics = null, bool verbose = false)
    {
        string pipeName = $"statistica-cli-{Guid.NewGuid():N}";
        using var pipe = new NamedPipeServerStream(pipeName, PipeDirection.InOut, 1, PipeTransmissionMode.Byte,
            PipeOptions.Asynchronous | PipeOptions.CurrentUserOnly);
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(timeoutSeconds));
        var startInfo = new ProcessStartInfo(executable)
        {
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardError = true,
            RedirectStandardOutput = true
        };
        startInfo.ArgumentList.Add("--internal-worker");
        startInfo.ArgumentList.Add(pipeName);

        long started = DateTime.UtcNow.Ticks;
        using var process = Process.Start(startInfo) ?? throw new InvalidOperationException("Could not start the COM worker.");
        int session = process.SessionId;
        var stdout = process.StandardOutput.ReadToEndAsync();
        var stderr = process.StandardError.ReadToEndAsync();
        using var reader = new StreamReader(pipe, leaveOpen: true);
        var events = new List<WorkerEvent>();
        var ended = new List<int>();
        var json = WorkerJson.Default;
        var pending = new Stack<WorkerEvent>();
        WorkerFailure? failure = null;
        WorkerResult? result = null;
        var records = new List<JobRecord>();
        try
        {
            var connect = pipe.WaitForConnectionAsync(timeout.Token);
            var exit = process.WaitForExitAsync(timeout.Token);
            if (await Task.WhenAny(connect, exit) == exit && !pipe.IsConnected)
            {
                await exit;
                throw new IOException("The worker exited before opening the diagnostic pipe.");
            }
            await connect;
            byte[] requestLine = Encoding.UTF8.GetBytes(JsonSerializer.Serialize(request, json.WorkerRequest) + "\n");
            await pipe.WriteAsync(requestLine, timeout.Token);
            while (await reader.ReadLineAsync(timeout.Token) is { } line)
            {
                var message = JsonSerializer.Deserialize(line, json.WorkerMessage)
                    ?? throw new InvalidDataException("The worker sent an empty message.");
                if (message.Event is { } entry)
                {
                    events.Add(entry);
                    if (entry.Event == "job_begin")
                    {
                        pending.Clear();
                        timeout.CancelAfter(TimeSpan.FromSeconds(timeoutSeconds));
                    }
                    if (entry.Event == "begin")
                        pending.Push(entry);
                    else if (entry.Event is "end" or "error" && pending.TryPeek(out var call) &&
                        call.Phase == entry.Phase && call.Operation == entry.Operation)
                        pending.Pop();
                    if (verbose)
                        Console.Error.WriteLine($"Job {entry.Job} {entry.Event} {entry.Phase}/{entry.Operation} {entry.Backend} {entry.Message}");
                    diagnostics?.WriteLine(JsonSerializer.Serialize(entry, json.WorkerEvent));
                }
                else if (message.Failure is { } reported)
                    failure = reported;
                else if (message.Record is { } checkpoint)
                    records.Add(checkpoint);
                else if (message.Result is { } completed)
                    result = completed;
            }
            await exit;
            if (failure is null && (process.ExitCode != 0 || (request.Jobs is null && result is null)))
            {
                var lastError = events.LastOrDefault(entry => entry.Event == "error");
                failure = new("failed", lastError?.Phase ?? "Worker", lastError?.Operation ?? "Execute",
                    lastError?.Backend, lastError?.Hresult, lastError?.Job ?? 0,
                    lastError?.Message ?? $"STATISTICA worker exited with code {process.ExitCode}. {ReadOutput(stderr)}".TrimEnd());
            }
        }
        catch (OperationCanceledException) when (timeout.IsCancellationRequested)
        {
            var last = pending.TryPeek(out var call) ? call : events.LastOrDefault(entry => entry.Event == "job_begin");
            failure = new("timed_out", last?.Phase ?? "Worker", last?.Operation ?? "Start-Process",
                last?.Backend, null, last?.Job ?? 0, $"STATISTICA worker exceeded {timeoutSeconds} seconds.");
        }
        catch (Exception error) when (error is IOException or JsonException)
        {
            failure = new("failed", "Worker", "Pipe", null, null, 0, $"{error.Message} {ReadOutput(stderr)}".TrimEnd());
        }
        finally
        {
            if (!process.HasExited)
            {
                try { process.Kill(); }
                catch (InvalidOperationException) when (process.HasExited) { }
                catch (Win32Exception) when (process.HasExited) { }
                process.WaitForExit(5000);
            }

            if (failure is not null)
                foreach (var entry in events)
                    if (entry is { Event: "server", Backend: "Application", Pid: int pid, Session: int serverSession, StartTicks: long ticks } &&
                        StaProcesses.StopOwned(new(pid, serverSession, ticks), request.BaselinePids, started, session))
                        ended.Add(pid);
        }
        bool inspectDesktop = failure is not null || events.Any(entry => entry is { Event: "error", Phase: "Activate" });
        var workerExit = new WorkerExit(DateTimeOffset.UtcNow, process.Id, failure,
            inspectDesktop ? StaProcesses.Snapshot() : [], ended.ToArray(), ReadOutput(stdout));
        diagnostics?.WriteLine(JsonSerializer.Serialize(workerExit, json.WorkerExit));
        return new(result, failure, events.ToArray(), workerExit, records.ToArray());
    }

    private static string? ReadOutput(Task<string> read) =>
        read.IsCompletedSuccessfully && read.Result.Trim() is { Length: > 0 } text ? text : null;
}
