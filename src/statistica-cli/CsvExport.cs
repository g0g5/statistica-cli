using System.Globalization;
using System.Text;
using Microsoft.VisualBasic.FileIO;

namespace StatisticaCli;

internal static class CsvExport
{
    private sealed class NumericRange(int start)
    {
        internal int Start { get; } = start;
        internal int Count { get; set; }
        internal object?[,]? Values { get; set; }
    }

    private sealed record ColumnSource(int Index, double MissingCode, NumericRange? Range);

    internal static void Write(Spreadsheet sheet, CsvExportOptions options)
    {
        int cases = sheet.Cases, variables = sheet.Variables;
        int last = options.LastRow == 0 ? cases : options.LastRow;
        if (options.FirstRow > last || last > cases)
            throw new InvalidDataException($"Row range must be within 1..{cases}.");
        if (variables < 1)
            throw new InvalidDataException("The spreadsheet has no variables to export.");

        int[] indices = options.Columns is { Length: > 0 } selected
            ? selected : Enumerable.Range(Spreadsheet.FirstIndex, variables).ToArray();
        var ranges = new List<NumericRange>();
        var sources = new Dictionary<int, ColumnSource>();
        foreach (int index in indices.Order())
        {
            if (index < Spreadsheet.FirstIndex || index > variables)
                throw new InvalidDataException($"Column '{index}' is outside 1..{variables}.");
            NumericRange? range = null;
            // The bulk API can decode text differently on legacy ANSI spreadsheets.
            if (!options.DisplayText && sheet.ColumnType(index) != Spreadsheet.TextColumn)
            {
                if (ranges.Count == 0 || (long)ranges[^1].Start + ranges[^1].Count != index)
                    ranges.Add(new(index));
                range = ranges[^1];
                range.Count++;
            }
            sources.Add(index, new(index, sheet.MissingCode(index), range));
        }

        ColumnSource[] columns = Array.ConvertAll(indices, index => sources[index]);
        string[] header = ["case_index", .. indices.Select(index => $"v{index.ToString(CultureInfo.InvariantCulture)}")];
        int blockRows = Math.Clamp(65_536 / indices.Length, 1, 1024);
        using (var writer = new StreamWriter(new FileStream(options.Stage, FileMode.CreateNew), new UTF8Encoding(false, true)))
        {
            WriteRow(writer, header);
            var fields = new string[header.Length];
            int start = options.FirstRow;
            while (true)
            {
                int count = Math.Min(blockRows, last - start + 1);
                foreach (var range in ranges)
                {
                    range.Values = sheet.ReadBlock(start, count, range.Start, range.Count).Cells;
                    if (range.Values.GetLength(0) != count || range.Values.GetLength(1) != range.Count)
                        throw new InvalidDataException("STATISTICA returned an unexpected data block size.");
                }
                for (int offset = 0; offset < count; offset++)
                {
                    int row = start + offset;
                    fields[0] = row.ToString(CultureInfo.InvariantCulture);
                    for (int c = 0; c < columns.Length; c++)
                    {
                        var column = columns[c];
                        var value = column.Range is { } range
                            ? range.Values![offset, column.Index - range.Start]
                            : sheet.ReadCell(row, column.Index, column.MissingCode, options.DisplayText);
                        fields[c + 1] = Spreadsheet.IsMissing(value, column.MissingCode)
                            ? "" : Convert.ToString(value, CultureInfo.InvariantCulture) ?? "";
                    }
                    WriteRow(writer, fields);
                }
                if (start + (count - 1) == last)
                    break;
                start += count;
            }
        }
        Verify(options.Stage, header, options.FirstRow, last);
    }

    private static void WriteRow(StreamWriter writer, string[] fields)
    {
        for (int i = 0; i < fields.Length; i++)
        {
            if (i > 0)
                writer.Write(',');
            writer.Write('"');
            foreach (char character in fields[i])
            {
                if (character == '"')
                    writer.Write('"');
                writer.Write(character);
            }
            writer.Write('"');
        }
        writer.WriteLine();
    }

    private static void Verify(string path, string[] header, int first, int last)
    {
        using var csv = new TextFieldParser(path, new UTF8Encoding(false, true), detectEncoding: false)
        {
            HasFieldsEnclosedInQuotes = true,
            TrimWhiteSpace = false
        };
        csv.SetDelimiters(",");
        if (csv.ReadFields() is not { } fields || !header.SequenceEqual(fields))
            throw new InvalidDataException("The CSV header did not round-trip.");
        for (long row = first; row <= last; row++)
        {
            if (csv.ReadFields() is not { } record || record.Length != header.Length ||
                record[0] != row.ToString(CultureInfo.InvariantCulture))
                throw new InvalidDataException($"The CSV record for case {row} did not round-trip.");
        }
        if (!csv.EndOfData)
            throw new InvalidDataException("The CSV contains unexpected extra records.");
    }
}
