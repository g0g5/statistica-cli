using System.Text.Json;
using System.Text.Json.Serialization;

namespace StatisticaCli;

internal enum Backend { Auto, Application, Library }

internal enum WorkerOperation
{
    Probe, Inspect, ExportCsv, VerifyAutomation, HoldServer, CreateSta, EditSta, ImportCsv, RunMacro,
    CreateGraph, ExportGraph, Screenshot, ListWorkbook, ExtractWorkbookItem, BuildWorkbook
}

internal sealed record WorkerRequest(
    Backend Backend, bool Detailed, int[] BaselinePids, string? ApplicationError, string? LibraryError,
    WorkerOperation Operation, Job[]? Jobs = null, bool ContinueOnError = false);

internal sealed record CsvExportOptions(string Stage, int[]? Columns, int FirstRow, int LastRow, bool DisplayText);

internal sealed record SpreadsheetColumn(int Index, string Name, int Type, double MissingCode);

internal sealed record SpreadsheetInfo(string Path, int Cases, int Variables, SpreadsheetColumn[] Columns);

internal sealed record WorkerResult(
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)] ProbeResult? Probe = null,
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)] SpreadsheetInfo? Spreadsheet = null,
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)] WorkbookInfo? Workbook = null);

internal sealed record Registration(string View, string? Clsid, string Kind, string? Server);

internal sealed record ProbeComponent(bool Available, string? Version, string? Error, int? Pid, Registration[] Registration);

internal sealed record ProbeHost(string? Pwsh, bool X64, string Apartment, int Session);

internal sealed record ProbeResult(
    string? Com, bool Available, string? Version, bool? ApplicationAvailable, string? ApplicationError)
{
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public Dictionary<string, ProbeComponent>? Components { get; init; }

    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public ProbeHost? Host { get; init; }
}

internal sealed record WorkerFailure(string Status, string Phase, string Operation, string? Backend,
    string? Hresult, int Job, string Message)
{
    internal string Format() =>
        $"[{Phase}/{Operation}{(Hresult is null ? "" : $" {Hresult}")}] {Message}";
}

internal sealed record ServerIdentity(int Pid, int Session, long StartTicks);

internal sealed record WorkerEvent(string Event, string Phase, string Operation, string? Backend)
{
    public DateTimeOffset Time { get; init; } = DateTimeOffset.UtcNow;
    public int Job { get; init; } = 1;
    public int WorkerPid { get; init; } = Environment.ProcessId;
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public long? ElapsedMs { get; init; }
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Hresult { get; init; }
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Message { get; init; }
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? Pid { get; init; }
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public int? Session { get; init; }
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public long? StartTicks { get; init; }
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? Apartment { get; init; }
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public bool? X64 { get; init; }
}

internal sealed record WorkerMessage(
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)] WorkerEvent? Event = null,
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)] WorkerResult? Result = null,
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)] WorkerFailure? Failure = null,
    [property: JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)] JobRecord? Record = null);

internal sealed record DesktopProcess(int Pid, int Session, DateTime Start, string? CommandLine);

internal sealed record WorkerExit(DateTimeOffset Time, int WorkerPid, WorkerFailure? Failure,
    DesktopProcess[] DesktopProcesses, int[] EndedServerPids, string? WorkerStdout)
{
    public string Event => "worker_exit";
}

[JsonSourceGenerationOptions(PropertyNamingPolicy = JsonKnownNamingPolicy.SnakeCaseLower)]
[JsonSerializable(typeof(WorkerRequest))]
[JsonSerializable(typeof(WorkerMessage))]
[JsonSerializable(typeof(WorkerExit))]
[JsonSerializable(typeof(BatchSummary))]
[JsonSerializable(typeof(PublishEvent))]
[JsonSerializable(typeof(JsonElement))]
[JsonSerializable(typeof(string))]
[JsonSerializable(typeof(bool))]
[JsonSerializable(typeof(int))]
[JsonSerializable(typeof(double))]
[JsonSerializable(typeof(int[]))]
[JsonSerializable(typeof(string[]))]
internal sealed partial class WorkerJson : JsonSerializerContext
{
    internal static WorkerJson Indented { get; } =
        new WorkerJson(new System.Text.Json.JsonSerializerOptions
        {
            PropertyNamingPolicy = System.Text.Json.JsonNamingPolicy.SnakeCaseLower,
            WriteIndented = true
        });
}
