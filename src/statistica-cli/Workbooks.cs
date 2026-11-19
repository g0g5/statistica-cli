using StatisticaCli.Automation;

namespace StatisticaCli;

internal static class Workbooks
{
    private const int ExtractCopy = 1, InsertLastChild = 1;
    private static readonly string[] Types = ["Folder", "Spreadsheet", "Graph", "Report", "Macro", "Other"];

    private static IEnumerable<(WorkbookEntry Entry, DispatchObject Item)> Traverse(DispatchObject book)
    {
        DispatchObject? item = book.Get<DispatchObject>("Root");
        try
        {
            int index = 0;
            while (item is not null)
            {
                var next = (DispatchObject?)book.GetValue("NextItem", item);
                item.Dispose();
                item = next;
                if (item is null)
                    yield break;
                int code = item.Get<int>("Type");
                yield return (new(++index, item.Get<string>("Name"), item.Get<string>("ItemPath"),
                    code >= 0 && code < Types.Length ? Types[code] : "Other", code), item);
            }
        }
        finally { item?.Dispose(); }
    }

    internal static WorkbookInfo List(DispatchObject app, string input, string reported)
    {
        using var document = Documents.Open(app, "Workbooks", input);
        return new(reported, Traverse(document.Value).Select(node => node.Entry).ToArray());
    }

    internal static void Extract(DispatchObject app, JobOptions options, string input, string stage)
    {
        string type = Type(stage);
        using (var book = Documents.Open(app, "Workbooks", input))
        {
            int index;
            if (options.ItemIndex is { } selected)
                index = selected;
            else
            {
                var matches = Traverse(book.Value).Where(node => node.Entry.Path == options.ItemPath).Select(node => node.Entry.Index).ToArray();
                if (matches.Length > 1)
                    throw new InvalidDataException($"ItemPath is ambiguous. Use one of these ItemIndex values: {string.Join(", ", matches)}.");
                index = matches.Length == 1 ? matches[0] : throw new InvalidDataException("No matching workbook item.");
            }

            bool found = false;
            foreach (var (entry, item) in Traverse(book.Value))
            {
                if (entry.Index != index)
                    continue;
                if (entry.Type != type)
                    throw new InvalidDataException($"Workbook item {index} is {entry.Type}, which cannot be saved as {Path.GetExtension(stage)}.");
                using var extracted = Documents.Own(item.Call<DispatchObject>("Extract", ExtractCopy));
                extracted.SaveAs(stage);
                found = true;
                break;
            }
            if (!found)
                throw new InvalidDataException("No matching workbook item.");
        }
        using var reopened = Documents.Open(app, type + "s", stage);
    }

    internal static void Build(DispatchObject app, string[] paths, string[] reported, string stage)
    {
        string[] names = reported.Select(path => Path.GetFileNameWithoutExtension(path)).ToArray();
        using (var book = Documents.New(app, "Workbooks"))
        {
            using var root = book.Value.Get<DispatchObject>("Root");
            for (int index = 0; index < paths.Length; index++)
            {
                using var item = book.Value.Call<DispatchObject>("InsertFile", paths[index], root, InsertLastChild);
                item.Set("Name", names[index]);
            }
            book.SaveAs(stage);
        }
        var entries = List(app, stage, stage).Items;
        if (entries.Length != paths.Length)
            throw new InvalidDataException("The reopened workbook has a different item count.");
        for (int index = 0; index < entries.Length; index++)
            if (entries[index].Name != names[index] || entries[index].Type != Type(paths[index]))
                throw new InvalidDataException($"Reopened workbook item {index + 1} differs from its input.");
    }

    private static string Type(string path) =>
        Path.GetExtension(path).Equals(".sta", StringComparison.OrdinalIgnoreCase) ? "Spreadsheet" : "Graph";
}
