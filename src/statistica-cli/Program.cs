using System.CommandLine;
using System.Runtime.Versioning;

[assembly: SupportedOSPlatform("windows")]

namespace StatisticaCli;

internal static class Program
{
    private static async Task<int> Main(string[] args)
    {
        if (args is ["--internal-worker", var pipeName])
        {
            int exitCode = 1;
            var sta = new Thread(() =>
            {
                try { exitCode = StaWorker.Run(pipeName); }
                catch (Exception error) { Console.Error.WriteLine(error.Message); }
            });
            sta.SetApartmentState(ApartmentState.STA);
            sta.Start();
            sta.Join();
            return exitCode;
        }

        var root = new RootCommand("Automate TIBCO STATISTICA on Windows.");
        var backend = new Option<Backend>("--backend")
        {
            Description = "Select the COM backend.", DefaultValueFactory = _ => Backend.Auto, Recursive = true
        };
        backend.Validators.Add(value =>
        {
            if (!Enum.IsDefined(value.GetValueOrDefault<Backend>()))
                value.AddError("--backend must be Auto, Application, or Library.");
        });
        var timeout = new Option<int>("--timeout-seconds")
        {
            Description = "Worker timeout in seconds.", DefaultValueFactory = _ => 120, Recursive = true
        };
        timeout.Validators.Add(value =>
        {
            if (value.GetValueOrDefault<int>() is < 1 or > 86400)
                value.AddError("--timeout-seconds must be between 1 and 86400.");
        });
        var diagnosticPath = new Option<string?>("--diagnostic-path")
        {
            Description = "Write diagnostics to a new .jsonl file.", Recursive = true
        };
        var verbose = new Option<bool>("--verbose") { Description = "Show COM calls on standard error.", Recursive = true };
        foreach (var option in new Option[] { backend, timeout, diagnosticPath, verbose })
            root.Options.Add(option);

        JobCommands.Add(root, parse => parse.GetValue(backend), async (parse, jobs, batch, keepGoing, manifest) =>
        {
            string? log = parse.GetValue(diagnosticPath);
            if (log is not null)
            {
                log = JobCommands.Resolve(log, Directory.GetCurrentDirectory(), "DiagnosticPath", [".jsonl"]);
                if (jobs.Any(job => new[] { job.Destination, job.Options.InputPath, job.Options.SpecPath, job.Options.MacroPath }
                        .Concat(job.Options.DocumentPaths ?? []).Contains(log, StringComparer.OrdinalIgnoreCase)) ||
                    string.Equals(log, manifest, StringComparison.OrdinalIgnoreCase))
                    throw new ArgumentException("DiagnosticPath must differ from all inputs and destinations.");
            }
            using var diagnostics = log is null ? null : new StreamWriter(
                new FileStream(log, FileMode.CreateNew, FileAccess.Write, FileShare.Read)) { AutoFlush = true };
            return await JobRunner.Run(jobs, batch, keepGoing, parse.GetValue(backend), parse.GetValue(timeout),
                diagnostics, parse.GetValue(verbose));
        });
        SkillInstaller.Add(root);
        return await root.Parse(args).InvokeAsync();
    }
}
