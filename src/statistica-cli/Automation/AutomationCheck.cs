namespace StatisticaCli.Automation;

internal static class AutomationCheck
{
    internal static void Verify(DispatchObject app)
    {
        using var sheets = app.Get<DispatchObject>("Spreadsheets");
        using var sheet = sheets.Call<DispatchObject>("New");
        try
        {
            int defaults = sheet.Get<int>("NumberOfVariables");
            sheet.Call("AddVariables", "Sample X", defaults, 1, 0, 0);
            sheet.Call("AddVariables", "Sample Text", defaults + 1, 1, 1, 20);
            sheet.Call("AddVariables", "Sample Y", defaults + 2, 1, 0, 0);
            if (defaults > 0)
                sheet.Call("DeleteVariables", 1, defaults);
            int rows = sheet.Get<int>("NumberOfCases");
            if (rows < 2)
                sheet.Call("AddCases", rows, 2 - rows);
            if (rows > 2)
                sheet.Call("DeleteCases", 3, rows);

            VerifyBlock(sheet, new(new object?[,] { { 1.25, "alpha", 10.0 }, { 2.75, "beta", 20.0 } }));
            VerifyBlock(sheet, new(new object?[,] { { 9.5, "gamma", 30.0 }, { 8.5, "delta", 40.0 } }, 1, 1));
            var optional = sheet.Call<DataBlock>("GetData", 1, 2, 1, 3, OptionalArgument.Missing);
            Require(Equals(optional.Cells[1, 1], "delta"), "GetData optional argument did not use its default.");

            sheet.Set("VariableName", "Label", 2);
            Require(sheet.Get<string>("VariableName", 2) == "Label", "Indexed property did not round-trip.");
            try
            {
                sheet.Get<int>("NumberOfVariables", 1);
                throw new InvalidDataException("IDispatch accepted an unexpected argument.");
            }
            catch (AutomationException error) when ((uint)error.HResult == 0x80020011 &&
                error.Message.Contains("IDispatch.Invoke(NumberOfVariables)", StringComparison.Ordinal)) { }
            try
            {
                sheet.Call("SetData", 1, 1, new DataBlock(new object?[,] { { 1.0, "a" }, { 2.0, "b" }, { 3.0, "c" } }));
                throw new InvalidDataException("SetData accepted a block larger than the spreadsheet.");
            }
            catch (AutomationException error) when ((uint)error.HResult == 0x80020009 &&
                error.Message.Contains("0x80070057", StringComparison.Ordinal)) { }

            VerifyPlot(app, sheet);
        }
        finally { sheet.Call("Close", false); }
    }

    private static void VerifyBlock(DispatchObject sheet, DataBlock expected)
    {
        try { sheet.Call("SetData", 1, 1, expected); }
        catch (AutomationException error)
        {
            throw new AutomationException($"SetData SAFEARRAY bounds ({expected.FirstRow}, {expected.FirstColumn})",
                error.HResult, error.Message);
        }
        var read = sheet.Call<DataBlock>("GetData", 1, 2, 1, 3, false);
        Require(read.FirstRow == 1 && read.FirstColumn == 1 &&
            read.Cells.GetLength(0) == 2 && read.Cells.GetLength(1) == 3,
            $"Unexpected SAFEARRAY bounds ({read.FirstRow}, {read.FirstColumn}) and dimensions " +
            $"{read.Cells.GetLength(0)}x{read.Cells.GetLength(1)}.");
        for (int row = 0; row < 2; row++)
            for (int column = 0; column < 3; column++)
                Require(Equals(expected.Cells[row, column], read.Cells[row, column]),
                    $"SAFEARRAY cell ({row + 1}, {column + 1}) differs: " +
                    $"expected {expected.Cells[row, column]}, received {read.Cells[row, column]}.");
    }

    private static void VerifyPlot(DispatchObject app, DispatchObject sheet)
    {
        const int scatterAnalysis = 11003;
        using var analysis = app.Call<DispatchObject>("Analysis", scatterAnalysis, sheet);
        try
        {
            using var dialog = analysis.Get<DispatchObject>("Dialog");
            dialog.Set("GraphType", 0);
            dialog.Set("FitType", 0);
            dialog.Set("Variables", "1 |3");
            using var graphs = dialog.Get<DispatchObject>("Graphs");
            using var graph = graphs.Get<DispatchObject>("Item", 1);
            try
            {
                using var plot = graph.Get<DispatchObject>("Plot", 1);
                bool previous = plot.Get<bool>("DisplayLine");
                plot.Set("DisplayLine", !previous);
                Require(plot.Get<bool>("DisplayLine") == !previous, "Plot.DisplayLine did not round-trip.");
            }
            finally { graph.Call("Close", false); }
        }
        finally { analysis.Call("Close"); }
    }

    private static void Require(bool condition, string message)
    {
        if (!condition)
            throw new InvalidDataException(message);
    }
}
