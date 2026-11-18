using Axyra.Sheets.Style;

namespace Axyra.Sheets.Examples.FeatureDemo;

/// <summary>
/// Sort and filter.
/// </summary>
/// <remarks>
/// Sorting rewrites the rows in place; <see cref="SortKey"/> stacks as many keys as you
/// need. Filtering hides rows rather than removing them — <c>AutoFilter.Apply()</c>
/// returns how many it hid, and the data is all still there.
/// <para>
/// AutoFilter custom criteria compare <b>exactly</b>. For begins-with or contains, use
/// <c>CellRange.AdvancedFilter</c>, which does honour <c>*</c> and <c>?</c>.
/// </para>
/// </remarks>
public sealed class SortFilterExample : IExample
{
    private const string Content = "sortFilter-orders-content.xlsx";
    private const int Rows = 24;

    public string Name => "feature-demo/sort-filter";

    public string Description => "Sort by one or several keys and hide rows with AutoFilter criteria";

    public void Run(string outputDirectory)
    {
        using (Workbook wb = Unsorted())
        {
            wb.Save(Path.Combine(outputDirectory, "sort-filter-unsorted.xlsx"));
        }
        using (Workbook wb = SortOne())
        {
            wb.Save(Path.Combine(outputDirectory, "sort-filter-sort-one.xlsx"));
        }
        using (Workbook wb = SortTwo())
        {
            wb.Save(Path.Combine(outputDirectory, "sort-filter-sort-two.xlsx"));
        }
        using (Workbook wb = FilterValue())
        {
            wb.Save(Path.Combine(outputDirectory, "sort-filter-filter-value.xlsx"));
        }
        using (Workbook wb = FilterComparison())
        {
            wb.Save(Path.Combine(outputDirectory, "sort-filter-filter-comparison.xlsx"));
        }
        Console.WriteLine("Wrote five sort and filter workbooks to " + outputDirectory);
    }

    /// <summary>The source data, in the order it was written.</summary>
    private static Workbook Unsorted()
    {
        Workbook wb = Orders();
        Sheet sheet = wb.Sheets[0];
        Note(sheet, "Straight from the writer — regions interleaved, revenue unordered.");
        return wb;
    }

    /// <summary>Revenue, highest first.</summary>
    private static Workbook SortOne()
    {
        Workbook wb = Orders();
        Sheet sheet = wb.Sheets[0];
        sheet.GetRange(1, 0, Rows, 3).Sort(3, ascending: false);
        Note(sheet, "sort(column 3, ascending=false) — highest revenue first.");
        return wb;
    }

    /// <summary>Region ascending, then revenue descending inside each region.</summary>
    private static Workbook SortTwo()
    {
        Workbook wb = Orders();
        Sheet sheet = wb.Sheets[0];
        sheet.GetRange(1, 0, Rows, 3).Sort(new SortKey(0), new SortKey(3, Ascending: false));
        Note(sheet, "sort(SortKey.asc(0), SortKey.desc(3)) — grouped, then ranked.");
        return wb;
    }

    /// <summary>Show only the rows whose region matches.</summary>
    private static Workbook FilterValue()
    {
        Workbook wb = Orders();
        Sheet sheet = wb.Sheets[0];

        sheet.AutoFilter.Range = sheet.GetRange(0, 0, Rows, 3);
        sheet.AutoFilter[0].FilterByValues("East");
        int hidden = sheet.AutoFilter.Apply();

        Note(sheet, "setColumnFilter(0, \"East\") hid " + hidden + " of " + Rows
                + " rows. The data is intact — only the row flags changed.");
        return wb;
    }

    /// <summary>Custom criteria compare exactly; here that is what we want.</summary>
    private static Workbook FilterComparison()
    {
        Workbook wb = Orders();
        Sheet sheet = wb.Sheets[0];

        sheet.AutoFilter.Range = sheet.GetRange(0, 0, Rows, 3);
        sheet.AutoFilter[3].FilterByCondition(FilterOperator.GreaterThan, "3000");
        int hidden = sheet.AutoFilter.Apply();

        Note(sheet, "setColumnCustomFilter(3, GREATER_THAN, \"3000\") hid "
                + hidden + " rows. Wildcards are not supported here — "
                + "applyAdvancedFilter handles those.");
        return wb;
    }

    private static void Note(Sheet sheet, string text)
    {
        int row = Rows + 2;
        sheet[row, 0].Value = text;
        sheet[row, 0].Style = new CellStyle { Font = new() { Italic = true } };
    }

    /// <summary>The order book: deliberately unsorted, with repeated regions, read from a packaged workbook.</summary>
    private static Workbook Orders() => Workbook.Open(DemoBooks.Read(Content), WorkbookFormat.Xlsx);
}
