using System.Globalization;
using System.Text;
using System.Text.Json;
using Microsoft.VisualBasic.FileIO;
using StatisticaCli.Automation;

namespace StatisticaCli;

internal static class SpreadsheetWrites
{
    private const int DoubleColumn = 0, TextColumn = Spreadsheet.TextColumn, LongColumn = 2, ByteColumn = 3;
    private sealed record Column(string Name, int Type, int Length, string? Formula);
    private sealed record Storage(int Type, int Length, object Missing);

    internal static void Create(DispatchObject app, JobOptions options, string stage, string? csv)
    {
        using var spec = JobCommands.ReadJson(options.SpecPath!);
        var root = spec.RootElement;
        if (csv is null)
            Allow(root, "columns", "rows");
        else
            Allow(root, "columns");
        var columns = Array(root, "columns").Select(ReadColumn).ToArray();
        if (columns.Length == 0)
            throw new InvalidDataException("Supply a nonempty columns array.");
        if (columns.Select(column => column.Name).Distinct(StringComparer.OrdinalIgnoreCase).Count() != columns.Length ||
            columns.Any(column => column.Formula is not null))
            throw new InvalidDataException("Column names must be unique. Use edit-sta to set formulas.");

        int blockRows = Math.Clamp(65536 / columns.Length, 1, 1024);
        using var expected = new FileStream(stage + ".verify", FileMode.CreateNew, FileAccess.ReadWrite,
            FileShare.None, 4096, FileOptions.DeleteOnClose | FileOptions.SequentialScan);
        using var writer = new BinaryWriter(expected, Encoding.UTF8, leaveOpen: true);
        int rows = 0;
        Storage[] storage;
        using (var document = Documents.New(app, "Spreadsheets"))
        {
            var sheet = document.Value;
            int defaults = sheet.Get<int>("NumberOfVariables");
            foreach (var column in columns)
                AddColumn(sheet, column);
            if (defaults > 0)
                sheet.Call("DeleteVariables", 1, defaults);
            storage = Enumerable.Range(1, columns.Length).Select(index => GetStorage(sheet, index)).ToArray();

            IEnumerable<object?[]> data = csv is null ? JsonRows(root, columns.Length) : CsvRows(csv, columns, options.MissingToken);
            foreach (var chunk in data.Chunk(blockRows))
            {
                var block = new object?[chunk.Length, columns.Length];
                for (int row = 0; row < chunk.Length; row++)
                    for (int column = 0; column < columns.Length; column++)
                    {
                        object cell = Normalize(chunk[row][column], storage[column], $"Cell ({rows + row + 1},{column + 1})", csv is not null);
                        block[row, column] = cell;
                        if (cell is string value)
                            writer.Write(value);
                        else
                            writer.Write((double)cell);
                    }
                int end = checked(rows + chunk.Length);
                int cases = sheet.Get<int>("NumberOfCases");
                if (cases < end)
                    sheet.Call("AddCases", cases, end - cases);
                sheet.Call("SetData", rows + 1, 1, new DataBlock(block));
                rows = end;
            }
            if (rows == 0)
                throw new InvalidDataException("Supply at least one data row.");
            int remaining = sheet.Get<int>("NumberOfCases");
            if (remaining > rows)
                sheet.Call("DeleteCases", rows + 1, remaining);
            document.SaveAs(stage);
        }

        writer.Flush();
        expected.Position = 0;
        using var reader = new BinaryReader(expected, Encoding.UTF8, leaveOpen: true);
        using var reopened = Documents.Open(app, "Spreadsheets", stage);
        VerifyDimensions(reopened.Value, rows, columns.Length);
        for (int index = 1; index <= columns.Length; index++)
            VerifyColumn(reopened.Value, index, columns[index - 1]);
        int verified = 0;
        while (verified < rows)
        {
            int count = Math.Min(blockRows, rows - verified);
            var data = reopened.Value.Call<DataBlock>("GetData", verified + 1, count, 1, columns.Length, false).Cells;
            for (int row = 0; row < count; row++)
                for (int column = 0; column < columns.Length; column++)
                {
                    object value;
                    if (storage[column].Type == TextColumn)
                        value = reader.ReadString();
                    else
                        value = reader.ReadDouble();
                    if (!EqualCell(data[row, column], value))
                        throw new InvalidDataException($"Reopened cell ({verified + row + 1},{column + 1}) differs from the specification.");
                }
            verified += count;
        }
    }

    internal static void Edit(DispatchObject app, JobOptions options, string input, string stage)
    {
        using var spec = JobCommands.ReadJson(options.SpecPath!);
        var root = spec.RootElement;
        Allow(root, "addColumns", "addRows", "cells", "formulas");
        var newColumns = Array(root, "addColumns").Select(ReadColumn).ToArray();
        var cells = Array(root, "cells");
        var formulaEntries = Array(root, "formulas");
        int addRows = root.TryGetProperty("addRows", out var amount) && amount.ValueKind != JsonValueKind.Null ? amount.GetInt32() : 0;
        if (addRows < 0 || (addRows == 0 && newColumns.Length == 0 && cells.Length == 0 && formulaEntries.Length == 0))
            throw new InvalidDataException("Supply a nonnegative addRows and at least one change.");

        Documents.Copy(input, stage);
        var added = new Dictionary<int, Column>();
        var formulas = new Dictionary<int, string>();
        var written = new Dictionary<(int Row, int Column), JsonElement>();
        var values = new Dictionary<int, DataBlock>();
        int rows, columns;
        using (var document = Documents.Open(app, "Spreadsheets", stage, writable: true))
        {
            document.RequireWritable();
            var sheet = document.Value;
            if (addRows > 0)
            {
                int before = sheet.Get<int>("NumberOfCases");
                int target = checked(before + addRows);
                sheet.Call("AddCases", before, addRows);
                if (sheet.Get<int>("NumberOfCases") != target)
                    throw new InvalidDataException("STATISTICA did not add the requested rows.");
            }

            var names = Enumerable.Range(1, sheet.Get<int>("NumberOfVariables")).Select(index => sheet.Get<string>("VariableName", index))
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
            foreach (var column in newColumns)
            {
                if (!names.Add(column.Name))
                    throw new InvalidDataException($"The column '{column.Name}' already exists.");
                int index = AddColumn(sheet, column);
                added.Add(index, column);
                if (column.Formula is { } formula)
                    formulas.Add(index, formula);
            }
            rows = sheet.Get<int>("NumberOfCases");
            columns = sheet.Get<int>("NumberOfVariables");
            var storage = new Dictionary<int, Storage>();
            foreach (var cell in cells)
            {
                int row = cell.GetProperty("row").GetInt32(), column = cell.GetProperty("column").GetInt32();
                Address(row, column, rows, columns);
                var value = cell.GetProperty("value");
                if (!written.TryAdd((row, column), value))
                    throw new InvalidDataException($"Duplicate cell edit ({row},{column}).");
                if (!storage.TryGetValue(column, out var definition))
                    storage.Add(column, definition = GetStorage(sheet, column));
                object normalized = Normalize(ReadValue(value), definition, $"Cell ({row},{column})", false);
                sheet.Call("SetData", row, column, new DataBlock(new object?[,] { { normalized } }));
            }
            foreach (var entry in formulaEntries)
            {
                int index = entry.GetProperty("column").GetInt32();
                Address(1, index, rows, columns);
                if (!formulas.TryAdd(index, entry.GetProperty("formula").GetString() ?? ""))
                    throw new InvalidDataException($"Duplicate formula for column {index}.");
            }
            foreach (var (index, formula) in formulas)
                SetFormula(sheet, index, formula);
            foreach (int index in formulas.Keys)
                values.Add(index, sheet.Call<DataBlock>("GetData", 1, rows, index, 1, false));
            document.Save();
        }

        using var reopened = Documents.Open(app, "Spreadsheets", stage);
        var actual = reopened.Value;
        VerifyDimensions(actual, rows, columns);
        foreach (var (index, column) in added)
            VerifyColumn(actual, index, column);
        foreach (var (address, value) in written)
            if (!formulas.ContainsKey(address.Column))
            {
                object? expected = ReadValue(value);
                VerifyCell(actual, address.Row, address.Column, expected is "" ? null : expected, 0);
            }
        foreach (var (index, formula) in formulas)
        {
            if (actual.Get<string>("VariableLongName", index) != formula)
                throw new InvalidDataException($"Reopened formula differs for column {index}.");
            var saved = actual.Call<DataBlock>("GetData", 1, rows, index, 1, false).Cells;
            for (int row = 0; row < rows; row++)
                if (!EqualCell(saved[row, 0], values[index].Cells[row, 0]))
                    throw new InvalidDataException($"Reopened formula differs for column {index}, row {row + 1}.");
        }
    }

    internal static object? ReadValue(JsonElement value) => value.ValueKind switch
    {
        JsonValueKind.Null => null,
        JsonValueKind.String => value.GetString(),
        JsonValueKind.Number when value.TryGetDouble(out double number) && double.IsFinite(number) => number,
        _ => throw new InvalidDataException("Cell values must be finite JSON numbers, strings, or null.")
    };

    internal static void VerifyCell(DispatchObject sheet, int row, int column, object? expected, double tolerance)
    {
        Address(row, column, sheet.Get<int>("NumberOfCases"), sheet.Get<int>("NumberOfVariables"));
        bool matches = expected switch
        {
            null => sheet.Get<bool>("MissingData", row, column),
            string text => sheet.Get<string>("Text", row, column) == text,
            double number => !sheet.Get<bool>("MissingData", row, column) &&
                Math.Abs(sheet.Get<double>("Value", row, column) - number) <= tolerance,
            _ => false
        };
        if (!matches)
            throw new InvalidDataException($"Cell ({row},{column}) differs from its expected value '{expected}'.");
    }

    private static void SetFormula(DispatchObject sheet, int index, string formula)
    {
        if (formula.Length < 2 || formula[0] != '=')
            throw new InvalidDataException($"Column {index} requires a formula starting with '='.");
        if (!sheet.Call<bool>("IsValidFormula", formula, VariantOutput.Value, VariantOutput.Value, VariantOutput.Value, VariantOutput.Value))
            throw new InvalidDataException($"Invalid formula for column {index}.");
        sheet.Set("VariableLongName", formula, index);
        sheet.Call("Recalculate", index);
    }

    private static bool EqualCell(object? actual, object? expected) =>
        expected is double number && actual is IConvertible and not (string or bool or DateTime)
            ? Convert.ToDouble(actual, CultureInfo.InvariantCulture) == number
            : Equals(actual, expected) || (actual is null && expected is "");

    private static void Allow(JsonElement value, params string[] names)
    {
        foreach (var item in value.EnumerateObject())
            if (!names.Contains(item.Name, StringComparer.Ordinal))
                throw new InvalidDataException($"Unknown spreadsheet specification property '{item.Name}'.");
    }

    private static JsonElement[] Array(JsonElement root, string name) =>
        root.TryGetProperty(name, out var value) ? value.EnumerateArray().ToArray() : [];

    private static Column ReadColumn(JsonElement value)
    {
        Allow(value, "name", "type", "length", "formula");
        string? name = value.GetProperty("name").GetString();
        if (string.IsNullOrWhiteSpace(name))
            throw new InvalidDataException("Each column requires a name.");
        int? length = value.TryGetProperty("length", out var size) && size.ValueKind != JsonValueKind.Null ? size.GetInt32() : null;
        string? formula = value.TryGetProperty("formula", out var expression) ? expression.GetString() : null;
        return value.GetProperty("type").GetString()?.ToLowerInvariant() switch
        {
            "double" when length is null => new(name, DoubleColumn, 0, formula),
            "text" when length > 0 => new(name, TextColumn, length.Value, formula),
            _ => throw new InvalidDataException($"Column '{name}' requires type double or text. Text requires a positive length.")
        };
    }

    private static int AddColumn(DispatchObject sheet, Column column)
    {
        int index = sheet.Get<int>("NumberOfVariables") + 1;
        sheet.Call("AddVariables", column.Name, index - 1, 1, column.Type, column.Length);
        VerifyColumn(sheet, index, column);
        return index;
    }

    private static void VerifyColumn(DispatchObject sheet, int index, Column column)
    {
        if (sheet.Get<string>("VariableName", index) != column.Name || sheet.Get<int>("VariableType", index) != column.Type)
            throw new InvalidDataException($"Column {index} differs from the requested name or type.");
    }

    private static void VerifyDimensions(DispatchObject sheet, int rows, int columns)
    {
        if (sheet.Get<int>("NumberOfCases") != rows || sheet.Get<int>("NumberOfVariables") != columns)
            throw new InvalidDataException("Reopened spreadsheet dimensions differ from the saved dimensions.");
    }

    private static Storage GetStorage(DispatchObject sheet, int index)
    {
        int type = sheet.Get<int>("VariableType", index);
        return type == TextColumn
            ? new(type, sheet.Get<int>("VariableTypeLength", index), "")
            : new(type, 0, sheet.Get<double>("VariableMissingData", index));
    }

    private static object Normalize(object? value, Storage storage, string address, bool csv)
    {
        if (value is null)
            return storage.Missing;
        if (storage.Type == TextColumn)
            return value is string text && (storage.Length <= 0 || text.Length <= storage.Length)
                ? text : throw new InvalidDataException($"{address} requires text within the column's length.");
        double number;
        if (csv)
        {
            if (!double.TryParse((string)value, NumberStyles.Float, CultureInfo.InvariantCulture, out number))
                throw new InvalidDataException($"{address} is not an invariant-culture number.");
        }
        else
            number = value is double typed ? typed : throw new InvalidDataException($"{address} requires a JSON number or null.");
        if (!double.IsFinite(number) || number == (double)storage.Missing ||
            (storage.Type is LongColumn or ByteColumn && number != Math.Truncate(number)) ||
            (storage.Type == LongColumn && (number < int.MinValue || number > int.MaxValue)) ||
            (storage.Type == ByteColumn && (number < byte.MinValue || number > byte.MaxValue)))
            throw new InvalidDataException($"{address} is outside the numeric range or equals the missing-value code. Use null for a missing value.");
        return number;
    }

    private static void Address(int row, int column, int rows, int columns)
    {
        if (row < 1 || row > rows || column < 1 || column > columns)
            throw new InvalidDataException($"Spreadsheet cell ({row},{column}) is out of range.");
    }

    private static IEnumerable<object?[]> JsonRows(JsonElement root, int columns)
    {
        int index = 0;
        foreach (var row in root.GetProperty("rows").EnumerateArray())
        {
            index++;
            if (row.ValueKind != JsonValueKind.Array || row.GetArrayLength() != columns)
                throw new InvalidDataException($"Data row {index} must contain {columns} fields.");
            yield return row.EnumerateArray().Select(ReadValue).ToArray();
        }
    }

    private static IEnumerable<object?[]> CsvRows(string path, Column[] columns, string? missing)
    {
        using var parser = new TextFieldParser(path, new UTF8Encoding(false, true), true)
        {
            HasFieldsEnclosedInQuotes = true,
            TrimWhiteSpace = false
        };
        parser.SetDelimiters(",");
        var header = parser.ReadFields();
        if (header is null || !header.SequenceEqual(columns.Select(column => column.Name)))
            throw new InvalidDataException("CSV header does not match the column specification.");
        int index = 0;
        while (!parser.EndOfData)
        {
            var fields = parser.ReadFields()!;
            index++;
            if (fields.Length != columns.Length)
                throw new InvalidDataException($"Data row {index} must contain {columns.Length} fields.");
            yield return fields.Select(field => field.Length == 0 || field == missing ? null : (object?)field).ToArray();
        }
    }
}
