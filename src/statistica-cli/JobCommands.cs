using System.CommandLine;
using System.Text.Json;
using System.Text.Json.Serialization.Metadata;

namespace StatisticaCli;

internal static class JobCommands
{
    private sealed record Field(string Key, Option Option, Func<ParseResult, JsonElement> Read, bool Required);
    private sealed record Definition(WorkerOperation Action, string Name, string Description, Field[] Fields,
        string[]? Input = null, string[]? Output = null);

    private static Field Value<T>(string key, string description, JsonTypeInfo<T> type, bool required = false)
    {
        string name = "--" + JsonNamingPolicy.KebabCaseLower.ConvertName(key);
        var option = new Option<T>(name) { Description = description, Required = required };
        if (typeof(T).IsArray)
        {
            option.Arity = ArgumentArity.OneOrMore;
            option.AllowMultipleArgumentsPerToken = true;
        }
        return new(key, option, parse => JsonSerializer.SerializeToElement(parse.GetValue(option)!, type), required);
    }

    private static Definition[] Definitions()
    {
        var json = WorkerJson.Default;
        Field Input() => Value("InputPath", "Path to the input file.", json.String, true);
        Field Output() => Value("OutputPath", "Path to the output file.", json.String, true);
        Field Spec() => Value("SpecPath", "Path to the JSON specification.", json.String, true);
        Field Force() => Value("Force", "Replace an existing target file.", json.Boolean);
        var verify = new Option<string>("--verify-cells") { Description = "JSON object or .json file mapping row,column to expected values." };
        return [
            new(WorkerOperation.Probe, "probe", "Check STATISTICA COM availability.",
                [Value("Detailed", "Report COM registration and host details.", json.Boolean)]),
            new(WorkerOperation.Inspect, "inspect", "Read STA dimensions and column metadata.", [Input()], [".sta"]),
            new(WorkerOperation.ExportCsv, "export-csv", "Export STA rows to UTF-8 CSV.",
                [Input(), Output(), Value("Columns", "Comma-separated one-based column numbers.", json.String),
                 Value("FirstRow", "First row (default 1).", json.Int32), Value("LastRow", "Last row, or zero for all remaining rows.", json.Int32),
                 Value("DisplayText", "Export displayed text.", json.Boolean), Force()], [".sta"], [".csv"]),
            new(WorkerOperation.CreateSta, "create-sta", "Create a spreadsheet from a JSON specification.",
                [Output(), Spec(), Force()], Output: [".sta"]),
            new(WorkerOperation.EditSta, "edit-sta", "Edit a copy of a spreadsheet.", [Input(), Output(), Spec(), Force()], [".sta"], [".sta"]),
            new(WorkerOperation.ImportCsv, "import-csv", "Import a UTF-8 CSV using a column specification.",
                [Input(), Output(), Spec(), Value("MissingToken", "Additional literal for missing fields.", json.String), Force()], [".csv"], [".sta"]),
            new(WorkerOperation.RunMacro, "run-macro", "Run an SVB macro on a copy of a spreadsheet.",
                [Input(), Output(), Value("MacroPath", "Path to the SVB macro.", json.String, true),
                 Value("PassInputAsScriptArgument", "Pass the spreadsheet to GetScriptArgument.", json.Boolean),
                 new("VerifyCells", verify, parse => ReadCellExpectations(parse.GetRequiredValue(verify)), false),
                 Value("Tolerance", "Absolute numeric tolerance (default 1e-8).", json.Double), Force()], [".sta"], [".sta"]),
            new(WorkerOperation.CreateGraph, "create-graph", "Create a line or scatter graph.",
                [Input(), Output(), Value("GraphType", "Line or Scatter.", json.String, true),
                 Value("XColumn", "One-based X column for scatter graphs.", json.Int32),
                 Value("YColumns", "One or more one-based Y columns.", json.Int32Array, true),
                 Value("ConnectPoints", "Connect scatter points with lines.", json.Boolean), Force()], [".sta"], [".stg"]),
            new(WorkerOperation.ExportGraph, "export-graph", "Export a graph as PNG or SVG.", [Input(), Output(), Force()], [".stg"], [".png", ".svg"]),
            new(WorkerOperation.Screenshot, "screenshot", "Capture a STATISTICA window as PNG.",
                [Input(), Output(), Value("CropToDocument", "Crop to the active document.", json.Boolean),
                 Value("CaptureTimeoutSeconds", "Wait for a stable frame (default 8 seconds).", json.Int32), Force()],
                [".sta", ".stg", ".stw"], [".png"]),
            new(WorkerOperation.ListWorkbook, "list-workbook", "List workbook items in document order.", [Input()], [".stw"]),
            new(WorkerOperation.ExtractWorkbookItem, "extract-workbook-item", "Extract a workbook spreadsheet or graph.",
                [Input(), Output(), Value("ItemIndex", "One-based item index.", json.Int32),
                 Value("ItemPath", "Exact workbook item path.", json.String), Force()], [".stw"], [".sta", ".stg"]),
            new(WorkerOperation.BuildWorkbook, "build-workbook", "Assemble a workbook from STA and STG files.",
                [Output(), Value("DocumentPaths", "One or more input STA or STG paths in order.", json.StringArray, true), Force()], Output: [".stw"])
        ];
    }

    internal static void Add(RootCommand root, Func<ParseResult, Backend> backend,
        Func<ParseResult, Job[], bool, bool, string?, Task<int>> run)
    {
        var definitions = Definitions();
        foreach (var definition in definitions)
        {
            var command = new Command(definition.Name, definition.Description);
            foreach (var field in definition.Fields)
                command.Options.Add(field.Option);
            command.SetAction((parse, _) => Guard(() =>
            {
                var values = new Dictionary<string, JsonElement>(StringComparer.OrdinalIgnoreCase);
                foreach (var field in definition.Fields)
                    if (parse.GetResult(field.Option) is { Implicit: false })
                        values.Add(field.Key, field.Read(parse));
                var job = Validate(definition, values, 1, Directory.GetCurrentDirectory(), NewOutputs(), backend(parse));
                return run(parse, [job], false, false, null);
            }));
            root.Subcommands.Add(command);
        }

        var batch = new Command("batch", "Run a JSON manifest of jobs in order.");
        var manifest = new Option<string>("--manifest-path") { Required = true, Description = "Path to the JSON job manifest." };
        var keepGoing = new Option<bool>("--continue-on-error") { Description = "Continue independent jobs after a failure." };
        batch.Options.Add(manifest);
        batch.Options.Add(keepGoing);
        batch.SetAction((parse, _) => Guard(() =>
        {
            string path = Resolve(parse.GetRequiredValue(manifest), Directory.GetCurrentDirectory(), "ManifestPath", [".json"]);
            using var document = ReadJson(path);
            var data = Properties(document.RootElement);
            if (!data.TryGetValue("jobs", out var entries) || entries.ValueKind != JsonValueKind.Array || entries.GetArrayLength() == 0)
                throw new ArgumentException("Manifest must contain a nonempty jobs array.");
            var outputs = NewOutputs();
            var jobs = new List<Job>();
            foreach (var entry in entries.EnumerateArray())
            {
                var values = Properties(entry);
                string? action = values.Remove("Action", out var name) ? name.GetString() : null;
                var definition = definitions.FirstOrDefault(item => string.Equals(item.Action.ToString(), action, StringComparison.OrdinalIgnoreCase))
                    ?? throw new ArgumentException($"Unsupported action '{action}' in job {jobs.Count + 1}.");
                jobs.Add(Validate(definition, values, jobs.Count + 1, Path.GetDirectoryName(path)!, outputs, backend(parse)));
            }
            return run(parse, jobs.ToArray(), true, parse.GetValue(keepGoing), path);
        }));
        root.Subcommands.Add(batch);
    }

    internal static async Task<int> Guard(Func<Task<int>> action)
    {
        try { return await action(); }
        catch (Exception error) when (error is IOException or UnauthorizedAccessException or ArgumentException or
            JsonException or InvalidOperationException or FormatException or System.ComponentModel.Win32Exception)
        {
            Console.Error.WriteLine(error.Message);
            return 1;
        }
    }

    private static Dictionary<string, Job> NewOutputs() => new(StringComparer.OrdinalIgnoreCase);
    private static Dictionary<string, JsonElement> Properties(JsonElement element) =>
        element.EnumerateObject().ToDictionary(item => item.Name, item => item.Value, StringComparer.OrdinalIgnoreCase);

    private static JsonElement ReadCellExpectations(string value)
    {
        using var document = value.TrimStart().StartsWith('{')
            ? JsonDocument.Parse(value)
            : ReadJson(Resolve(value, Directory.GetCurrentDirectory(), "VerifyCells", [".json"]));
        return document.RootElement.Clone();
    }

    internal static JsonDocument ReadJson(string path) => JsonDocument.Parse(File.ReadAllText(path),
        new JsonDocumentOptions { CommentHandling = JsonCommentHandling.Skip });

    internal static string Resolve(string? path, string directory, string label, string[] extensions)
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException($"Supply {label}.");
        string fullPath = Path.GetFullPath(path, directory);
        if (!extensions.Contains(Path.GetExtension(fullPath), StringComparer.OrdinalIgnoreCase))
            throw new ArgumentException($"{label} requires {string.Join(" or ", extensions)}.");
        return fullPath;
    }

    private static Job Validate(Definition command, Dictionary<string, JsonElement> values, int id, string directory,
        Dictionary<string, Job> prior, Backend backend)
    {
        foreach (string key in values.Keys)
            if (!command.Fields.Any(field => string.Equals(field.Key, key, StringComparison.OrdinalIgnoreCase)))
                throw new ArgumentException($"Unknown option '{key}' for job {id} ({command.Action}).");
        foreach (var field in command.Fields)
            if (field.Required && !values.ContainsKey(field.Key))
                throw new ArgumentException($"Supply {field.Key} for job {id}.");

        string? String(string key) => values.TryGetValue(key, out var value) ? value.GetString() : null;
        bool Flag(string key) => values.TryGetValue(key, out var value) && value.GetBoolean();
        int? Integer(string key) => values.TryGetValue(key, out var value) ? value.GetInt32() : null;

        var inputs = new List<string>();
        string Input(string? path, string key, string[] extensions, bool allowDependency = false)
        {
            string resolved = Resolve(path, directory, key, extensions);
            if ((!allowDependency || !prior.ContainsKey(resolved)) && !File.Exists(resolved))
                throw new FileNotFoundException($"Input does not exist: {resolved}");
            if (allowDependency)
                inputs.Add(resolved);
            return resolved;
        }

        string? input = command.Input is { } inputExtensions ? Input(String("InputPath"), "InputPath", inputExtensions, true) : null;
        string? spec = values.ContainsKey("SpecPath") ? Input(String("SpecPath"), "SpecPath", [".json"]) : null;
        string? macro = values.ContainsKey("MacroPath") ? Input(String("MacroPath"), "MacroPath", [".svb"]) : null;
        string[]? documents = null;
        if (values.TryGetValue("DocumentPaths", out var paths))
        {
            documents = paths.ValueKind == JsonValueKind.String
                ? [Input(paths.GetString(), "DocumentPaths", [".sta", ".stg"], true)]
                : paths.EnumerateArray().Select(path => Input(path.GetString(), "DocumentPaths", [".sta", ".stg"], true)).ToArray();
            if (documents.Length == 0)
                throw new ArgumentException("Supply at least one DocumentPaths entry.");
        }

        int first = Integer("FirstRow") ?? 1, last = Integer("LastRow") ?? 0;
        int capture = Integer("CaptureTimeoutSeconds") ?? 8;
        if (first < 1 || last < 0 || capture is < 1 or > 120)
            throw new ArgumentException("FirstRow, LastRow or CaptureTimeoutSeconds is outside its supported range.");
        foreach (string name in new[] { "Columns", "MissingToken" })
            if (values.TryGetValue(name, out var value) && value.ValueKind != JsonValueKind.String)
                throw new ArgumentException($"{name} must be a string.");
        int[]? columns = String("Columns")?.Split(',').Select(item =>
            int.TryParse(item, out int index) && index > 0 ? index : throw new ArgumentException($"Invalid column number: '{item}'."))
            .Distinct().ToArray();
        string? graphType = String("GraphType");
        int? x = Integer("XColumn");
        bool connect = Flag("ConnectPoints");
        int[]? y = null;
        if (command.Action == WorkerOperation.CreateGraph)
        {
            var series = values["YColumns"];
            y = series.ValueKind == JsonValueKind.Number ? [series.GetInt32()] : series.EnumerateArray().Select(item => item.GetInt32()).ToArray();
            if (y.Length == 0 || y.Any(index => index < 1) || y.Distinct().Count() != y.Length)
                throw new ArgumentException("YColumns requires distinct positive column numbers.");
            if (string.Equals(graphType, "Line", StringComparison.OrdinalIgnoreCase) && x is null && !connect)
                graphType = "Line";
            else if (string.Equals(graphType, "Scatter", StringComparison.OrdinalIgnoreCase) && x is > 0 && !y.Contains(x.Value))
                graphType = "Scatter";
            else
                throw new ArgumentException("Supply Line or Scatter. Scatter requires a positive XColumn distinct from YColumns. ConnectPoints applies to Scatter.");
        }

        int? itemIndex = Integer("ItemIndex");
        string? itemPath = String("ItemPath");
        if (command.Action == WorkerOperation.ExtractWorkbookItem &&
            (itemIndex.HasValue == !string.IsNullOrWhiteSpace(itemPath) || itemIndex is < 1))
            throw new ArgumentException("Supply exactly one positive ItemIndex or one ItemPath.");
        double tolerance = values.TryGetValue("Tolerance", out var tol) ? tol.GetDouble() : 1e-8;
        if (!double.IsFinite(tolerance) || tolerance < 0)
            throw new ArgumentException("Tolerance must be finite and nonnegative.");
        var expectations = values.TryGetValue("VerifyCells", out var cells)
            ? cells.EnumerateObject().ToDictionary(item => item.Name, item => item.Value.Clone(), StringComparer.Ordinal)
            : null;
        bool pass = Flag("PassInputAsScriptArgument");
        bool desktop = command.Action == WorkerOperation.Screenshot || (command.Action == WorkerOperation.RunMacro && !pass);
        if (desktop && backend == Backend.Library)
            throw new ArgumentException($"{command.Action} requires the Application backend in this mode.");

        string? output = null, stage = null;
        bool force = Flag("Force");
        if (command.Output is { } outputExtensions)
        {
            output = Resolve(String("OutputPath"), directory, "OutputPath", outputExtensions);
            string parent = Path.GetDirectoryName(output)!;
            if (!Directory.Exists(parent))
                throw new DirectoryNotFoundException($"Output directory does not exist: {parent}");
            if (inputs.Contains(output, StringComparer.OrdinalIgnoreCase) || prior.ContainsKey(output))
                throw new ArgumentException($"Output must differ from its inputs and prior outputs: {output}");
            if (Directory.Exists(output) || (File.Exists(output) && !force))
                throw new IOException($"Output already exists: {output}");
            stage = Path.Combine(parent, $".statistica-{Guid.NewGuid():N}{Path.GetExtension(output)}");
        }
        var options = new JobOptions(input, spec, macro, String("MissingToken"), Flag("Detailed"), columns, first, last,
            Flag("DisplayText"), pass, expectations, tolerance, graphType, x, y, connect, Flag("CropToDocument"), capture,
            itemIndex, itemPath, documents);
        var job = new Job(id, command.Action, options, output, stage, force, desktop,
            inputs.Where(prior.ContainsKey).Select(path => prior[path].Id).Distinct().Order().ToArray());
        if (output is not null)
            prior.Add(output, job);
        return job;
    }
}
