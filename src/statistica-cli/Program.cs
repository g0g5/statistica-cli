using System.CommandLine;
using System.ComponentModel;
using System.Runtime.Versioning;
using System.Text.Json;
using System.Text.Json.Serialization.Metadata;

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

        async Task<int> Run(ParseResult parse, WorkerOperation operation, bool detailed = false, string? input = null,
            string? output = null, int[]? columns = null, int firstRow = 1, int lastRow = 0,
            bool displayText = false, bool force = false)
        {
            string? stage = null;
            try
            {
                if (input is not null)
                {
                    input = ResolvePath(input, ".sta", "--input-path");
                    if (!File.Exists(input))
                        throw new FileNotFoundException($"Input spreadsheet does not exist: {input}");
                }
                CsvExportOptions? export = null;
                if (output is not null)
                {
                    output = ResolvePath(output, ".csv", "--output-path");
                    string directory = Path.GetDirectoryName(output)!;
                    if (!Directory.Exists(directory))
                        throw new DirectoryNotFoundException($"Output directory does not exist: {directory}");
                    if (Directory.Exists(output) || (File.Exists(output) && !force))
                        throw new IOException($"Output already exists: {output}");
                    stage = Path.Combine(directory, $".statistica-{Guid.NewGuid():N}.csv");
                    export = new(stage, columns, firstRow, lastRow, displayText);
                }
                var request = new WorkerRequest(parse.GetValue(backend), detailed, StaProcesses.Baseline(),
                    null, null, operation, input, export);
                using var diagnostics = parse.GetValue(diagnosticPath) is { } requestedLog ? new StreamWriter(
                    new FileStream(ResolvePath(requestedLog, ".jsonl", "--diagnostic-path"),
                        FileMode.CreateNew, FileAccess.Write, FileShare.Read)) { AutoFlush = true } : null;
                var run = await WorkerRunner.RunWithFallbackAsync(Environment.ProcessPath!, request,
                    parse.GetValue(timeout), diagnostics, parse.GetValue(verbose));
                if (run.Failure is { } failure)
                {
                    Console.Error.WriteLine(failure.Format());
                    return 1;
                }

                switch (operation)
                {
                    case WorkerOperation.Probe:
                        WriteJson(run.Result!.Probe!, WorkerJson.Indented.ProbeResult);
                        break;
                    case WorkerOperation.Inspect:
                        WriteJson(run.Result!.Spreadsheet!, WorkerJson.Indented.SpreadsheetInfo);
                        break;
                    case WorkerOperation.ExportCsv:
                        try { File.Move(stage!, output!, overwrite: force); }
                        catch (IOException error)
                        {
                            Console.Error.WriteLine($"[Publish/File.Move] {error.Message}");
                            return 1;
                        }
                        Console.WriteLine(output);
                        break;
                }
                return 0;
            }
            catch (Exception error) when (error is IOException or UnauthorizedAccessException or Win32Exception or ArgumentException)
            {
                Console.Error.WriteLine(error.Message);
                return 1;
            }
            finally
            {
                if (stage is not null)
                {
                    try { File.Delete(stage); }
                    catch (Exception error) when (error is IOException or UnauthorizedAccessException)
                    {
                        Console.Error.WriteLine($"Could not remove temporary output {stage}: {error.Message}");
                    }
                }
            }
        }

        var probe = new Command("probe", "Check STATISTICA COM availability.");
        var detailed = new Option<bool>("--detailed") { Description = "Report COM registration and host details." };
        probe.Options.Add(detailed);
        probe.SetAction((parse, _) => Run(parse, WorkerOperation.Probe, detailed: parse.GetValue(detailed)));
        root.Subcommands.Add(probe);

        var inspect = new Command("inspect", "Read row counts and column metadata from an STA spreadsheet.");
        var inspectInput = InputOption();
        inspect.Options.Add(inspectInput);
        inspect.SetAction((parse, _) => Run(parse, WorkerOperation.Inspect, input: parse.GetRequiredValue(inspectInput)));
        root.Subcommands.Add(inspect);

        var csv = new Command("export-csv", "Export STA rows to a UTF-8 CSV file.");
        var csvInput = InputOption();
        var destination = new Option<string>("--output-path") { Description = "Path to the CSV file.", Required = true };
        var selection = new Option<int[]?>("--columns")
        {
            Description = "Comma-separated one-based column numbers.",
            Arity = ArgumentArity.ExactlyOne,
            CustomParser = result =>
            {
                var selected = new HashSet<int>();
                var columns = new List<int>();
                foreach (string text in result.Tokens[0].Value.Split(','))
                {
                    if (!int.TryParse(text, out int index) || index < Spreadsheet.FirstIndex)
                    {
                        result.AddError($"Invalid column number: '{text}'.");
                        return null;
                    }
                    if (selected.Add(index))
                        columns.Add(index);
                }
                return columns.ToArray();
            }
        };
        var first = RowOption("--first-row", "First one-based row to export.", Spreadsheet.FirstIndex);
        var last = RowOption("--last-row", "Last row to export. Zero means the final row.", 0);
        var display = new Option<bool>("--display-text") { Description = "Export the displayed cell text." };
        var overwrite = new Option<bool>("--force") { Description = "Replace an existing target file." };
        foreach (var option in new Option[] { csvInput, destination, selection, first, last, display, overwrite })
            csv.Options.Add(option);
        csv.SetAction((parse, _) => Run(parse, WorkerOperation.ExportCsv, input: parse.GetRequiredValue(csvInput),
            output: parse.GetRequiredValue(destination), columns: parse.GetValue(selection), firstRow: parse.GetValue(first),
            lastRow: parse.GetValue(last), displayText: parse.GetValue(display), force: parse.GetValue(overwrite)));
        root.Subcommands.Add(csv);

        return await root.Parse(args).InvokeAsync();
    }

    private static Option<string> InputOption() => new("--input-path")
    {
        Description = "Path to the STA spreadsheet.", Required = true
    };

    private static Option<int> RowOption(string name, string description, int minimum)
    {
        var option = new Option<int>(name) { Description = description, DefaultValueFactory = _ => minimum };
        option.Validators.Add(result =>
        {
            if (result.GetValueOrDefault<int>() < minimum)
                result.AddError($"{name} must be at least {minimum}.");
        });
        return option;
    }

    private static string ResolvePath(string path, string extension, string option)
    {
        string fullPath = Path.GetFullPath(path);
        if (!string.Equals(Path.GetExtension(fullPath), extension, StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException($"{option} must be a file with extension {extension}.");
        return fullPath;
    }

    private static void WriteJson<T>(T result, JsonTypeInfo<T> type)
    {
        using var output = Console.OpenStandardOutput();
        JsonSerializer.Serialize(output, result, type);
        output.WriteByte((byte)'\n');
    }
}
