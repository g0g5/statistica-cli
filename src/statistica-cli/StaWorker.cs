using System.Diagnostics;
using System.IO.Pipes;
using System.Runtime.InteropServices;
using System.Text.Json;
using Microsoft.Win32;
using StatisticaCli.Automation;

namespace StatisticaCli;

internal sealed partial class StaWorker(WorkerRequest request, Action<WorkerMessage> send)
{
    private sealed class CallFailed(WorkerFailure failure) : Exception(failure.Message)
    {
        internal WorkerFailure Failure { get; } = failure;
    }

    private readonly Dictionary<Backend, string> activationErrors = new();
    private Backend backend;

    internal static int Run(string pipeName)
    {
        using var pipe = new NamedPipeClientStream(".", pipeName, PipeDirection.InOut);
        pipe.Connect(10_000);
        using var reader = new StreamReader(pipe, leaveOpen: true);
        using var writer = new StreamWriter(pipe, leaveOpen: true) { AutoFlush = true };
        var json = WorkerJson.Default;
        var request = JsonSerializer.Deserialize(reader.ReadLine() ?? "", json.WorkerRequest)
            ?? throw new InvalidDataException("Missing worker request.");

        void Send(WorkerMessage message) => writer.WriteLine(JsonSerializer.Serialize(message, json.WorkerMessage));
        var worker = new StaWorker(request, Send);
        if (request.Jobs is { } jobs)
            return worker.RunJobs(jobs);
        worker.Emit("job_begin", "Job", request.Operation.ToString());
        try
        {
            WorkerResult result;
            using (var apartment = new AutomationApartment())
                result = worker.Execute(apartment, request.Detailed);
            worker.Emit("job_end", "Job", request.Operation.ToString());
            Send(new(Result: result));
            return 0;
        }
        catch (Exception error)
        {
            var failure = error is CallFailed known ? known.Failure : worker.Report("Worker", "Execute", error);
            Send(new(Failure: failure));
            return 1;
        }
    }

    private WorkerResult Execute(AutomationApartment apartment, bool detailed)
    {
        if (request.ApplicationError is { } applicationError)
            activationErrors.TryAdd(Backend.Application, applicationError);
        if (request.LibraryError is { } libraryError)
            activationErrors.TryAdd(Backend.Library, libraryError);

        var candidates = request.Backend == Backend.Auto ? [Backend.Application, Backend.Library] : new[] { request.Backend };
        var components = new Dictionary<string, ProbeComponent>();
        ProbeResult? result = null;
        foreach (Backend candidate in candidates)
        {
            backend = candidate;
            string? version = null;
            int? pid = null;
            if (!activationErrors.ContainsKey(backend))
            {
                DispatchObject? app = null;
                try
                {
                    (app, pid) = Activate(apartment);
                }
                catch (CallFailed error)
                {
                    activationErrors[backend] = error.Failure.Format();
                    if (!detailed && candidates.Length == 1)
                        throw;
                }
                if (app is not null)
                {
                    using (app)
                    {
                        try
                        {
                            try { version = Call("Probe", "Application.Version", () => app.Get<string>("Version")); }
                            catch (CallFailed error) when (detailed && request.Operation == WorkerOperation.Probe)
                            {
                                activationErrors[backend] = error.Failure.Format();
                            }
                            if (version is not null)
                            {
                                if (request.Operation == WorkerOperation.VerifyAutomation)
                                    Call("Verify", "Spreadsheet Automation", () => { AutomationCheck.Verify(app); return true; });
                                if (request.Operation == WorkerOperation.HoldServer)
                                    Call("Verify", "HoldServer", () => { Thread.Sleep(Timeout.Infinite); return true; });
                            }
                        }
                        finally
                        {
                            if (backend == Backend.Application)
                                Call("Cleanup", "Application.Quit", () => { app.Call("Quit"); return true; });
                        }
                    }
                }
            }

            string? errorText = activationErrors.GetValueOrDefault(backend);
            if (version is not null)
                result ??= new($"Statistica.{backend}", true, version, ApplicationAvailable(backend),
                    activationErrors.GetValueOrDefault(Backend.Application));
            if (!detailed && result is not null)
                return new(Probe: result);
            if (detailed)
                components.Add(backend.ToString(), new(version is not null, version, errorText,
                    version is null ? null : pid, Registrations(backend)));
        }
        if (!detailed)
            throw new CallFailed(Report("Activate", "Open-StaSession",
                new InvalidOperationException("STATISTICA COM activation failed. " + string.Join(" | ", activationErrors.Values))));

        result ??= new(null, false, null, ApplicationAvailable(null), activationErrors.GetValueOrDefault(Backend.Application));
        using var process = Process.GetCurrentProcess();
        return new(Probe: result with
        {
            Components = components,
            Host = new(null, Environment.Is64BitProcess, "STA", process.SessionId)
        });
    }

    private WorkerResult ExecuteSpreadsheet(DispatchObject app, string input, WorkerOperation operation,
        string? reportedPath, CsvExportOptions? export)
    {
        using var sheet = Call("Open", "Spreadsheets.OpenReadOnly", () => Spreadsheet.OpenReadOnly(app, input));
        try
        {
            if (operation == WorkerOperation.Inspect)
                return new(Spreadsheet: Call("Read", "Spreadsheet variable metadata", () => sheet.Inspect(reportedPath ?? input)));

            Call("Export", "Spreadsheet data to CSV", () =>
            {
                CsvExport.Write(sheet, export!);
                return true;
            });
            return new();
        }
        finally { Call("Cleanup", "Spreadsheet.Close", () => { sheet.Close(); return true; }); }
    }

    private bool? ApplicationAvailable(Backend? selected) =>
        selected == Backend.Application ? true : activationErrors.ContainsKey(Backend.Application) ? false : null;

    private (DispatchObject App, int Pid) Activate(AutomationApartment apartment)
    {
        using var process = Process.GetCurrentProcess();
        long started = process.StartTime.ToUniversalTime().Ticks;
        int[] beforeActivation = backend == Backend.Application ? StaProcesses.Baseline() : [];
        var app = Call("Activate", "CoCreateInstance", () =>
            DispatchObject.Activate(apartment, $"Statistica.{backend}", backend == Backend.Application));
        try
        {
            var identity = Call("Activate", "Application.ProcessID", () =>
            {
                int pid = app.Get<int>("ProcessID");
                var server = StaProcesses.Identify(pid);
                if (backend == Backend.Application &&
                    (request.BaselinePids.Contains(pid) || beforeActivation.Contains(pid) ||
                        server.StartTicks < started || server.Session != process.SessionId))
                    throw new InvalidOperationException($"COM returned a pre-existing or external STATISTICA process ({pid}).");
                if (backend == Backend.Library && pid != Environment.ProcessId)
                    throw new InvalidOperationException($"The Library backend returned an unexpected server PID ({pid}).");
                return server;
            });
            Emit("server", "Activate", "Application.ProcessID", entry => entry with
            {
                Pid = identity.Pid, Session = identity.Session, StartTicks = identity.StartTicks,
                Apartment = "STA", X64 = Environment.Is64BitProcess
            });
            return (app, identity.Pid);
        }
        catch
        {
            app.Dispose();
            throw;
        }
    }

    private T Call<T>(string phase, string operation, Func<T> body)
    {
        var watch = Stopwatch.StartNew();
        Emit("begin", phase, operation);
        try
        {
            var result = body();
            Emit("end", phase, operation, entry => entry with { ElapsedMs = watch.ElapsedMilliseconds });
            return result;
        }
        catch (CallFailed)
        {
            Emit("end", phase, operation, entry => entry with { ElapsedMs = watch.ElapsedMilliseconds });
            throw;
        }
        catch (DocumentCleanupException error)
        {
            Emit("end", phase, operation, entry => entry with { ElapsedMs = watch.ElapsedMilliseconds });
            throw new CallFailed(Report("Cleanup", "Document", error, watch.ElapsedMilliseconds));
        }
        catch (Exception error) { throw new CallFailed(Report(phase, operation, error, watch.ElapsedMilliseconds)); }
    }

    private WorkerFailure Report(string phase, string operation, Exception error, long? elapsedMs = null)
    {
        string? hresult = error is COMException ? $"0x{error.HResult:X8}" : null;
        Emit("error", phase, operation, entry => entry with
        {
            Hresult = hresult, Message = error.Message, ElapsedMs = elapsedMs
        });
        return new("failed", phase, operation, backend == Backend.Auto ? null : backend.ToString(), hresult, jobId, error.Message);
    }

    private void Emit(string type, string phase, string operation, Func<WorkerEvent, WorkerEvent>? enrich = null)
    {
        var entry = new WorkerEvent(type, phase, operation, backend == Backend.Auto ? null : backend.ToString()) { Job = jobId };
        send(new(Event: enrich?.Invoke(entry) ?? entry));
    }

    private static Registration[] Registrations(Backend backend)
    {
        var entries = new List<Registration>();
        foreach (var view in new[] { RegistryView.Registry64, RegistryView.Registry32 })
        {
            using var root = RegistryKey.OpenBaseKey(RegistryHive.ClassesRoot, view);
            using var key = root.OpenSubKey($@"Statistica.{backend}\CLSID");
            if (key is null)
                continue;
            string? id = key.GetValue("") as string;
            string kind = backend == Backend.Application ? "LocalServer32" : "InprocServer32";
            using var server = root.OpenSubKey($@"CLSID\{id}\{kind}");
            entries.Add(new(view.ToString(), id, kind, server?.GetValue("") as string));
        }
        return entries.ToArray();
    }
}
