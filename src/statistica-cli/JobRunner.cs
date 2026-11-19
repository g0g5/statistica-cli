using System.Text.Json;

namespace StatisticaCli;

internal static class JobRunner
{
    internal static async Task<int> Run(Job[] jobs, bool batch, bool keepGoing, Backend backend, int timeout,
        TextWriter? diagnostics, bool verbose)
    {
        try
        {
            bool detailed = jobs[0].Options.Detailed;
            var request = new WorkerRequest(backend, detailed, StaProcesses.Baseline(), null, null,
                jobs[0].Action, Jobs: jobs, ContinueOnError: keepGoing);
            var run = await WorkerRunner.RunWithFallbackAsync(Environment.ProcessPath!, request, timeout, diagnostics, verbose);
            var records = run.Records.ToDictionary(record => record.Id);
            var entries = new List<BatchEntry>();
            bool stopped = false;
            foreach (var job in jobs)
            {
                var entry = new BatchEntry(job.Id, job.Action.ToString(), "skipped", null, null, null, null);
                var blocked = job.Dependencies.Where(id => entries[id - 1].Status != "succeeded").ToArray();
                if (blocked.Length > 0)
                    entry = entry with { Error = $"Dependency jobs did not publish successfully: {string.Join(", ", blocked)}." };
                else if (stopped)
                    entry = entry with { Error = "A previous job failed." };
                else if (records.TryGetValue(job.Id, out var record) && (batch || run.Failure is null))
                    entry = Publish(job, entry, record);
                else if (run.Failure is { } failure && (!batch || failure.Job == job.Id || (failure.Job == 0 && job.Id == 1)))
                    entry = entry with { Status = failure.Status, Backend = failure.Backend, Error = failure.Format() };
                else
                    entry = entry with { Error = "The worker did not complete this job." };

                entries.Add(entry);
                stopped |= !keepGoing && entry.Status != "succeeded";
                if (job.Destination is not null)
                    diagnostics?.WriteLine(JsonSerializer.Serialize(new PublishEvent(DateTimeOffset.UtcNow, job.Id,
                        entry.Status, job.Destination, entry.Error), WorkerJson.Default.PublishEvent));
            }
            if (batch)
            {
                Console.WriteLine(JsonSerializer.Serialize(new BatchSummary(entries.ToArray(), run.Failure), WorkerJson.Indented.BatchSummary));
                int incomplete = entries.Count(entry => entry.Status != "succeeded");
                if (incomplete > 0 || run.Failure is not null)
                {
                    Console.Error.WriteLine($"STATISTICA batch did not finish successfully ({incomplete} incomplete jobs).");
                    return 1;
                }
                return 0;
            }
            var single = entries[0];
            if (single.Error is { } error)
            {
                Console.Error.WriteLine(error);
                return 1;
            }
            if (single.Output is { } output)
                Console.WriteLine(output);
            else if (single.Result is { } result)
                Console.WriteLine(JsonSerializer.Serialize(result, WorkerJson.Indented.JsonElement));
            return 0;
        }
        finally
        {
            foreach (var job in jobs)
            {
                if (job.Stage is not { } stage)
                    continue;
                try { File.Delete(stage); }
                catch (Exception error) when (error is IOException or UnauthorizedAccessException)
                {
                    Console.Error.WriteLine($"Could not remove temporary output {stage}: {error.Message}");
                }
            }
        }
    }

    private static BatchEntry Publish(Job job, BatchEntry entry, JobRecord record)
    {
        entry = entry with { Backend = record.Backend };
        if (record.Failure is { } failure)
            return entry with { Status = "failed", Error = failure.Format() };
        try
        {
            if (job.Destination is { } destination)
                File.Move(job.Stage!, destination, job.Force);
            var json = WorkerJson.Default;
            JsonElement? result = record.Result switch
            {
                { Probe: { } probe } => JsonSerializer.SerializeToElement(probe, json.ProbeResult),
                { Spreadsheet: { } sheet } => JsonSerializer.SerializeToElement(sheet, json.SpreadsheetInfo),
                { Workbook: { } book } => JsonSerializer.SerializeToElement(book, json.WorkbookInfo),
                _ => null
            };
            return entry with { Status = "succeeded", Output = job.Destination, Result = result };
        }
        catch (Exception error) when (error is IOException or UnauthorizedAccessException)
        {
            return entry with { Status = "failed", Error = $"[Publish/File.Move] {error.Message}" };
        }
    }
}

internal sealed record PublishEvent(DateTimeOffset Time, int Job, string Status, string Destination, string? Message)
{
    public string Event => "publish";
    public string Phase => "Publish";
    public string Operation => "File.Move";
}
