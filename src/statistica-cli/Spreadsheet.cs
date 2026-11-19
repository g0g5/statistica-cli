using StatisticaCli.Automation;

namespace StatisticaCli;

internal sealed class Spreadsheet : IDisposable
{
    internal const int FirstIndex = 1;
    internal const int TextColumn = 1;
    private readonly DispatchObject sheet;

    private Spreadsheet(DispatchObject sheet) => this.sheet = sheet;

    internal static Spreadsheet OpenReadOnly(DispatchObject app, string path)
    {
        using var documents = app.Get<DispatchObject>("Spreadsheets");
        return new(documents.Call<DispatchObject>("OpenReadOnly", path, false));
    }

    internal int Cases => sheet.Get<int>("NumberOfCases");
    internal int Variables => sheet.Get<int>("NumberOfVariables");
    internal int ColumnType(int index) => sheet.Get<int>("VariableType", index);
    internal double MissingCode(int index) => sheet.Get<double>("VariableMissingData", index);

    internal SpreadsheetInfo Inspect(string path)
    {
        int variables = Variables;
        var columns = new SpreadsheetColumn[variables];
        for (int i = 0; i < columns.Length; i++)
        {
            int index = i + FirstIndex;
            columns[i] = new(index, sheet.Get<string>("VariableName", index), ColumnType(index), MissingCode(index));
        }
        return new(path, Cases, variables, columns);
    }

    internal DataBlock ReadBlock(int firstRow, int rowCount, int firstColumn, int columnCount) =>
        sheet.Call<DataBlock>("GetData", firstRow, rowCount, firstColumn, columnCount, false);

    internal object? ReadCell(int row, int column, double missingCode, bool displayText)
    {
        using var cell = sheet.Get<DispatchObject>("Cells", row, column);
        object? value = cell.GetValue("Value");
        if (IsMissing(value, missingCode))
            return null;
        return displayText ? cell.Get<string>("Text") : value;
    }

    internal static bool IsMissing(object? value, double code) =>
        value is null || (value is IConvertible and not (string or bool or DateTime) && Convert.ToDouble(value) == code);

    internal void Close() => sheet.Call("Close", false);

    public void Dispose() => sheet.Dispose();
}
