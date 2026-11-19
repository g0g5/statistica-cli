using StatisticaCli.Automation;

namespace StatisticaCli;

internal static class Graphs
{
    private const int LineAnalysis = 11012, ScatterAnalysis = 11003;
    private const int SingleY = 0, MultipleY = 1, NoFit = 0, PngFormat = 6;

    internal static void Create(DispatchObject app, JobOptions options, string input, string stage)
    {
        int[] series = options.YColumns!;
        using (var spreadsheet = Documents.Open(app, "Spreadsheets", input))
        {
            var sheet = spreadsheet.Value;
            int variables = sheet.Get<int>("NumberOfVariables");
            var indices = options.XColumn is { } x ? series.Append(x) : series;
            foreach (int index in indices)
            {
                if (index > variables)
                    throw new InvalidDataException($"Graph column {index} is outside 1..{variables}.");
                if (sheet.Get<int>("VariableType", index) == Spreadsheet.TextColumn)
                    throw new InvalidDataException($"Graph column {index} must be numeric.");
            }

            bool scatter = options.GraphType == "Scatter";
            using var analysis = app.Call<DispatchObject>("Analysis", scatter ? ScatterAnalysis : LineAnalysis, sheet);
            try
            {
                using var dialog = analysis.Get<DispatchObject>("Dialog");
                dialog.Set("GraphType", series.Length > 1 ? MultipleY : SingleY);
                dialog.Set("FitType", NoFit);
                string columns = string.Join(" ", series);
                dialog.Set("Variables", scatter ? $"{options.XColumn} |{columns}" : columns);
                using var graphs = dialog.Get<DispatchObject>("Graphs");
                int count = graphs.Get<int>("Count");
                if (count != 1)
                    throw new InvalidDataException($"Expected one graph, received {count}.");
                using var document = Documents.Own(graphs.Get<DispatchObject>("Item", 1));
                VerifyPlots(document.Value, series.Length);
                if (options.ConnectPoints)
                    for (int index = 1; index <= series.Length; index++)
                    {
                        using var plot = document.Value.Get<DispatchObject>("Plot", index);
                        plot.Set("DisplayLine", true);
                    }
                document.SaveAs(stage);
            }
            finally { analysis.Call("Close"); }
        }

        using var reopened = Documents.Open(app, "Graphs", stage);
        VerifyPlots(reopened.Value, series.Length);
        if (options.ConnectPoints)
            for (int index = 1; index <= series.Length; index++)
            {
                using var plot = reopened.Value.Get<DispatchObject>("Plot", index);
                if (!plot.Get<bool>("DisplayLine"))
                    throw new InvalidDataException($"Reopened plot {index} has no connecting line.");
            }
    }

    private static void VerifyPlots(DispatchObject graph, int expected)
    {
        int actual = graph.Get<int>("PlotsCount");
        if (actual != expected)
            throw new InvalidDataException($"Expected {expected} plots, received {actual}.");
    }

    internal static void Export(DispatchObject app, string input, string stage)
    {
        using var document = Documents.Open(app, "Graphs", input);
        if (Path.GetExtension(stage).Equals(".png", StringComparison.OrdinalIgnoreCase))
        {
            if (!document.Value.Call<bool>("SaveAsFormat", stage, PngFormat, true))
                throw new IOException("PNG export returned false.");
        }
        else
            document.Value.Call("ExportSVG", stage, 0, 0, 0);
    }
}
