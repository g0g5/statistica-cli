using StatisticaCli.Automation;

namespace StatisticaCli;

internal sealed partial class StaWorker
{
    private int jobId = 1;
    private DispatchObject? session;
    private int sessionPid;

    private int RunJobs(Job[] jobs)
    {
        var records = new Dictionary<int, JobRecord>();
        if (request.ApplicationError is { } applicationError)
            activationErrors[Backend.Application] = applicationError;
        if (request.LibraryError is { } libraryError)
            activationErrors[Backend.Library] = libraryError;

        try
        {
            using var apartment = new AutomationApartment();
            try
            {
                foreach (var job in jobs)
                {
                    jobId = job.Id;
                    string operation = job.Action.ToString();
                    Emit("job_begin", "Job", operation);
                    var failed = job.Dependencies.Where(id => records[id].Failure is not null).ToArray();
                    WorkerResult? result = null;
                    WorkerFailure? failure = null;
                    string? usedBackend = null;
                    try
                    {
                        if (failed.Length > 0)
                            throw new CallFailed(Report("Dependency", operation,
                                new InvalidOperationException($"Dependency jobs failed: {string.Join(", ", failed)}.")));
                        var options = job.Options;
                        string Substitute(string path)
                        {
                            foreach (int id in job.Dependencies)
                                if (string.Equals(path, jobs[id - 1].Destination, StringComparison.OrdinalIgnoreCase))
                                    return jobs[id - 1].Stage!;
                            return path;
                        }
                        if (job.Action == WorkerOperation.Probe && options.Detailed)
                        {
                            CloseSession();
                            result = Execute(apartment, true);
                            usedBackend = result.Probe?.Com?.Split('.').Last();
                        }
                        else
                        {
                            if (job.NeedsDesktop)
                                CloseSession();
                            var app = OpenSession(apartment, job);
                            usedBackend = backend.ToString();
                            result = ExecuteJob(app, job, options.InputPath is { } path ? Substitute(path) : null,
                                options.DocumentPaths?.Select(Substitute).ToArray());
                            if (job.NeedsDesktop)
                            {
                                CloseSession();
                                if (job.Action == WorkerOperation.RunMacro)
                                {
                                    var verifier = OpenSession(apartment, job);
                                    Call("Verify", "Reopened macro output", () => Macros.Verify(verifier, options, job.Stage!));
                                    CloseSession();
                                }
                            }
                        }
                        if (job.Stage is { } stage)
                            Call("Verify", "Output file", () => Documents.VerifyOutput(stage));
                    }
                    catch (Exception error)
                    {
                        failure = error switch
                        {
                            CallFailed known => known.Failure,
                            DocumentCleanupException => Report("Cleanup", "Document", error),
                            _ => Report("Job", operation, error)
                        };
                        if (job.NeedsDesktop)
                            CloseSession();
                    }
                    var record = new JobRecord(job.Id, usedBackend ?? failure?.Backend, result, failure);
                    records.Add(job.Id, record);
                    send(new(Record: record));
                    Emit("job_end", "Job", operation);
                    if (failure is not null && (!request.ContinueOnError || failure.Phase == "Cleanup"))
                        break;
                }
            }
            finally { CloseSession(); }
            return 0;
        }
        catch (Exception error)
        {
            send(new(Failure: error is CallFailed known ? known.Failure : Report("Worker", "Execute", error)));
            return 1;
        }
    }

    private DispatchObject OpenSession(AutomationApartment apartment, Job job)
    {
        if (session is not null)
            return session;

        Backend[] candidates = job.NeedsDesktop ? [Backend.Application]
            : request.Backend != Backend.Auto ? [request.Backend]
            : job.Action is WorkerOperation.Probe or WorkerOperation.Inspect or WorkerOperation.ExportCsv or WorkerOperation.RunMacro
                ? [Backend.Application, Backend.Library] : [Backend.Library, Backend.Application];
        CallFailed? last = null;
        foreach (var candidate in candidates)
        {
            backend = candidate;
            if (!job.NeedsDesktop && activationErrors.ContainsKey(candidate))
                continue;
            try
            {
                (session, sessionPid) = Activate(apartment);
                return session;
            }
            catch (CallFailed error)
            {
                last = error;
                activationErrors[candidate] = error.Failure.Format();
            }
        }

        if (last is not null && candidates.Length == 1)
            throw last;
        throw new CallFailed(Report("Activate", "Open-StaSession",
            new InvalidOperationException("STATISTICA COM activation failed. " + string.Join(" | ", activationErrors.Values))));
    }

    private void CloseSession()
    {
        if (session is not { } app)
            return;
        session = null;
        using (app)
            if (backend == Backend.Application)
                Call("Cleanup", "Application.Quit", () => app.Call("Quit"));
    }

    private void Call(string phase, string operation, Action action) =>
        Call(phase, operation, () => { action(); return true; });

    private WorkerResult ExecuteJob(DispatchObject app, Job job, string? input, string[]? paths)
    {
        var options = job.Options;
        string stage = job.Stage!;
        switch (job.Action)
        {
            case WorkerOperation.Probe:
                return new(Probe: new($"Statistica.{backend}", true,
                    Call("Probe", "Application.Version", () => app.Get<string>("Version")), ApplicationAvailable(backend),
                    activationErrors.GetValueOrDefault(Backend.Application)));
            case WorkerOperation.Inspect:
            case WorkerOperation.ExportCsv:
                return ExecuteSpreadsheet(app, input!, job.Action, options.InputPath,
                    job.Action == WorkerOperation.ExportCsv
                        ? new(stage, options.Columns, options.FirstRow, options.LastRow, options.DisplayText) : null);
            case WorkerOperation.CreateSta:
            case WorkerOperation.ImportCsv:
                Call("Write", "Create spreadsheet", () => SpreadsheetWrites.Create(app, options, stage, input));
                break;
            case WorkerOperation.EditSta:
                Call("Write", "Edit spreadsheet", () => SpreadsheetWrites.Edit(app, options, input!, stage));
                break;
            case WorkerOperation.RunMacro:
                Macros.Run(app, input!, options, stage, job.NeedsDesktop, Call);
                if (!job.NeedsDesktop)
                    Call("Verify", "Reopened macro output", () => Macros.Verify(app, options, stage));
                break;
            case WorkerOperation.CreateGraph:
                Call("Graph", "Create graph", () => Graphs.Create(app, options, input!, stage));
                break;
            case WorkerOperation.ExportGraph:
                Call("Export", "Export graph", () => Graphs.Export(app, input!, stage));
                break;
            case WorkerOperation.ListWorkbook:
                return new(Workbook: Call("Read", "Workbook.NextItem", () => Workbooks.List(app, input!, options.InputPath!)));
            case WorkerOperation.ExtractWorkbookItem:
                Call("Extract", "WorkbookItem.Extract(copy)", () => Workbooks.Extract(app, options, input!, stage));
                break;
            case WorkerOperation.BuildWorkbook:
                Call("Write", "Workbook.InsertFile", () => Workbooks.Build(app, paths!, options.DocumentPaths!, stage));
                break;
            case WorkerOperation.Screenshot:
                Call("Capture", "Capture window", () => WindowCapture.Save(app, sessionPid, options, input!, stage, Call));
                break;
            default:
                throw new InvalidDataException($"Unsupported operation: {job.Action}.");
        }
        return new();
    }
}
