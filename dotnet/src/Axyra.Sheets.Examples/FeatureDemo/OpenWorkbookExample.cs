using Axyra.Sheets.Style;

namespace Axyra.Sheets.Examples.FeatureDemo;

/// <summary>
/// Open an existing workbook — from a stream, from bytes, and with a look at what the reader
/// found.
/// </summary>
/// <remarks>
/// Opening is <b>lazy</b>: nothing is recalculated until you ask for it, so a file with stale
/// cached results keeps them until <c>Recalculate()</c> runs.
/// </remarks>
public sealed class OpenWorkbookExample : IExample
{
    private const string Book = "budget.xlsx";

    public string Name => "feature-demo/open-workbook";

    public string Description => "Open a workbook from a stream or bytes, inspect it, and recalculate on demand";

    public void Run(string outputDirectory)
    {
        using (Workbook wb = OpenFromStream())
        {
            wb.Save(Path.Combine(outputDirectory, "open-workbook-from-stream.xlsx"));
        }
        using (Workbook wb = OpenFromBytes())
        {
            wb.Save(Path.Combine(outputDirectory, "open-workbook-from-bytes.xlsx"));
        }
        using (Workbook wb = Inspect())
        {
            wb.Save(Path.Combine(outputDirectory, "open-workbook-inspect.xlsx"));
        }
        using (Workbook wb = OpenAndRecalculate())
        {
            wb.Save(Path.Combine(outputDirectory, "open-workbook-recalculated.xlsx"));
        }
        Console.WriteLine("Opened " + Book + " four ways; results in " + outputDirectory);
    }

    /// <summary>The reader is chosen from the format, so any source works.</summary>
    private static Workbook OpenFromStream()
    {
        using MemoryStream stream = new(DemoBooks.Read(Book));
        return Workbook.Open(stream, WorkbookFormat.Xlsx);
    }

    /// <summary>Straight from a byte[] — no temporary file, no stream to close.</summary>
    private static Workbook OpenFromBytes() => Workbook.Open(DemoBooks.Read(Book), WorkbookFormat.Xlsx);

    /// <summary>
    /// What the reader found: sheet names, indexes and used extents, written onto a sheet of
    /// their own.
    /// </summary>
    private static Workbook Inspect()
    {
        Workbook wb = OpenFromBytes();
        Sheet summary = wb.Sheets.Add("Opened");
        summary["A1:D1"].SetValues(new CellValue[,] { { "Index", "Sheet", "Used range", "Cells" } });
        summary["A1:D1"].Style = new CellStyle { Font = new Font { Bold = true } };

        int row = 1;
        for (int i = 0; i < wb.Sheets.Count; i++)
        {
            Sheet sheet = wb.Sheets[i];
            if (sheet.Index == summary.Index)
            {
                continue;
            }
            CellRange? used = sheet.UsedRange;
            summary[row, 0].Value = sheet.Index;
            summary[row, 1].Value = sheet.Name;
            summary[row, 2].Value = used == null ? "(empty)" : Address(used);
            summary[row, 3].Value = used == null ? 0 : (long)used.RowCount * used.ColumnCount;
            row++;
        }
        summary.Columns[1].Width = 24;
        summary.Columns[2].Width = 16;
        wb.ActiveSheet = summary;
        return wb;
    }

    /// <summary>Cached results survive the open; recalculation is opt-in.</summary>
    private static Workbook OpenAndRecalculate()
    {
        Workbook wb = OpenFromBytes();
        wb.Recalculate();
        return wb;
    }

    /// <summary>Formats a range as an A1 address such as <c>A1:D12</c>.</summary>
    private static string Address(CellRange r) =>
        Col(r.FirstColumn) + (r.FirstRow + 1) + ":" + Col(r.LastColumn) + (r.LastRow + 1);

    private static string Col(int index)
    {
        string name = "";
        for (int c = index + 1; c > 0; c = (c - 1) / 26)
        {
            name = (char)('A' + (c - 1) % 26) + name;
        }
        return name;
    }
}
