using System.Text.Json;

namespace StatisticaCli;

internal sealed record JobOptions(
    string? InputPath = null, string? SpecPath = null, string? MacroPath = null, string? MissingToken = null,
    bool Detailed = false, int[]? Columns = null, int FirstRow = 1, int LastRow = 0, bool DisplayText = false,
    bool PassInputAsScriptArgument = false, Dictionary<string, JsonElement>? VerifyCells = null, double Tolerance = 1e-8,
    string? GraphType = null, int? XColumn = null, int[]? YColumns = null, bool ConnectPoints = false,
    bool CropToDocument = false, int CaptureTimeoutSeconds = 8, int? ItemIndex = null, string? ItemPath = null,
    string[]? DocumentPaths = null);

internal sealed record Job(int Id, WorkerOperation Action, JobOptions Options, string? Destination, string? Stage,
    bool Force, bool NeedsDesktop, int[] Dependencies);

internal sealed record JobRecord(int Id, string? Backend, WorkerResult? Result, WorkerFailure? Failure);
internal sealed record BatchEntry(int Id, string Action, string Status, string? Backend, string? Output,
    JsonElement? Result, string? Error);
internal sealed record BatchSummary(BatchEntry[] Jobs, WorkerFailure? WorkerError);

internal sealed record WorkbookEntry(int Index, string Name, string Path, string Type, int TypeCode);
internal sealed record WorkbookInfo(string Path, WorkbookEntry[] Items);
