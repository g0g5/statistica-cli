using System.Globalization;
using StatisticaCli.Automation;

namespace StatisticaCli;

internal static class Macros
{
    private static (int Row, int Column, object? Value)[] Expectations(JobOptions options) =>
        options.VerifyCells?.Select(cell =>
        {
            var parts = cell.Key.Split(',');
            if (parts.Length != 2 || !Index(parts[0], out int row) || !Index(parts[1], out int column))
                throw new InvalidDataException($"Invalid VerifyCells address '{cell.Key}'. Use row,column with 1-based indices.");
            return (row, column, SpreadsheetWrites.ReadValue(cell.Value));
        }).ToArray() ?? [];

    private static bool Index(string text, out int index) =>
        int.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out index) && index > 0 && text[0] != '0';

    internal static void Run(DispatchObject app, string input, JobOptions options, string stage, bool visible,
        Action<string, string, Action> call)
    {
        _ = Expectations(options);
        call("Stage", "File.Copy", () => Documents.Copy(input, stage));
        bool openedMacro = false;
        try
        {
            using var document = Documents.Open(app, "Spreadsheets", stage, writable: true);
            document.LeaveForQuit = visible;
            document.RequireWritable();
            if (visible)
                document.Value.Set("Visible", true);
            using var macro = Documents.Open(app, "Macros", options.MacroPath!);
            openedMacro = true;
            macro.LeaveForQuit = visible;
            if (visible)
            {
                call("Open", "Spreadsheet.Activate", () => document.Value.Call("Activate"));
                call("Execute", "Macro.Execute", () => macro.Value.Call("Execute"));
            }
            else
                call("Execute", "Macro.ExecuteWithArgument", () => macro.Value.Call("ExecuteWithArgument", document.Value));
            call("Save", "Spreadsheet.Save", document.Save);
        }
        finally
        {
            if (openedMacro)
                call("Cleanup", "Macro documents", () => ClearDocuments(app, visible));
        }
    }

    private static void ClearDocuments(DispatchObject app, bool visible)
    {
        foreach (string name in new[] { "Macros", "Graphs", "Spreadsheets", "Workbooks", "Reports" })
        {
            using var collection = app.Get<DispatchObject>(name);
            for (int index = collection.Get<int>("Count"); index >= 1; index--)
            {
                using var document = Documents.Own(collection.Get<DispatchObject>("Item", index));
                document.LeaveForQuit = visible;
            }
        }
    }

    internal static void Verify(DispatchObject app, JobOptions options, string stage)
    {
        using var reopened = Documents.Open(app, "Spreadsheets", stage);
        _ = reopened.Value.Get<int>("NumberOfCases");
        _ = reopened.Value.Get<int>("NumberOfVariables");
        foreach (var (row, column, value) in Expectations(options))
            SpreadsheetWrites.VerifyCell(reopened.Value, row, column, value, options.Tolerance);
    }
}
