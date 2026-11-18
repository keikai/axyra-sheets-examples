using Axyra.Sheets.Style;

namespace Axyra.Sheets.Examples.FeatureDemo;

/// <summary>
/// Find and replace.
/// </summary>
/// <remarks>
/// <c>FindAll</c> returns <b>every</b> hit as a list of cells — there is no cursor to
/// advance and no "find next" to loop. <see cref="FindOptions"/> covers case sensitivity,
/// whole-cell matching and <c>*</c>/<c>?</c> wildcards, and regular expressions for the
/// harder patterns.
/// </remarks>
public sealed class FindReplaceExample : IExample
{
    private const string Content = "findReplace-tickets-content.xlsx";

    public string Name => "feature-demo/find-replace";

    public string Description => "Find every match with options or a regex, and replace in one call";

    public void Run(string outputDirectory)
    {
        using (Workbook wb = FindAndHighlight())
        {
            wb.Save(Path.Combine(outputDirectory, "find-replace-highlight.xlsx"));
        }
        using (Workbook wb = FindWithOptions())
        {
            wb.Save(Path.Combine(outputDirectory, "find-replace-options.xlsx"));
        }
        using (Workbook wb = ReplaceAll())
        {
            wb.Save(Path.Combine(outputDirectory, "find-replace-replace-all.xlsx"));
        }
        using (Workbook wb = RegexReplace())
        {
            wb.Save(Path.Combine(outputDirectory, "find-replace-regex.xlsx"));
        }
        Console.WriteLine("Wrote four find and replace workbooks to " + outputDirectory);
    }

    /// <summary>Every hit comes back at once, so highlighting them is a single loop.</summary>
    private static Workbook FindAndHighlight()
    {
        Workbook wb = Tickets();
        Sheet sheet = wb.Sheets[0];

        IReadOnlyList<Cell> hits = sheet.FindAll("overdue");
        var flag = new CellStyle
        {
            Fill = Color.FromRgb(0xFF, 0xC7, 0xCE),
            Font = new() { Color = Color.FromRgb(0x9C, 0x00, 0x06), Bold = true },
        };
        foreach (Cell hit in hits)
        {
            hit.Style = flag;
        }

        Note(sheet, "find(\"overdue\") returned " + hits.Count + " hits — all of them, "
                + "in one call.");
        return wb;
    }

    /// <summary>Case, whole-cell and wildcard switches, counted side by side.</summary>
    private static Workbook FindWithOptions()
    {
        Workbook wb = Tickets();
        Sheet sheet = wb.Sheets[0];

        int @base = sheet.UsedRange!.LastRow + 3;
        sheet.GetRange(@base, 0, @base, 1).SetValues(new CellValue[,] { { "Call", "Hits" } });
        sheet.GetRange(@base, 0, @base, 1).Style = new CellStyle { Font = new() { Bold = true } };

        Count(sheet, @base + 1, "find(\"overdue\")",
                sheet.FindAll("overdue").Count);
        Count(sheet, @base + 2, "find(\"OVERDUE\", caseSensitive=true, ...)",
                sheet.FindAll("OVERDUE", new FindOptions { MatchCase = true }).Count);
        Count(sheet, @base + 3, "find(\"overdue\", ..., matchEntire=true, ...)",
                sheet.FindAll("overdue", new FindOptions { MatchEntireCell = true }).Count);
        Count(sheet, @base + 4, "find(\"over*\", ..., wildcard=true)",
                sheet.FindAll("over*", new FindOptions { UseWildcards = true }).Count);
        Count(sheet, @base + 5, "findRegex(\"^INV-\\\\d{4}$\")",
                sheet.FindAll(@"^INV-\d{4}$", new FindOptions { UseRegex = true }).Count);

        sheet.Columns[0].Width = 42;
        return wb;
    }

    /// <summary><c>ReplaceAll</c> returns how many cells it changed.</summary>
    private static Workbook ReplaceAll()
    {
        Workbook wb = Tickets();
        Sheet sheet = wb.Sheets[0];

        int changed = sheet.ReplaceAll("overdue", "ESCALATED");
        Note(sheet, "replaceAll(\"overdue\", \"ESCALATED\") changed " + changed + " cells.");

        var flag = new CellStyle
        {
            Fill = Color.FromRgb(0xFF, 0xEB, 0x9C),
            Font = new() { Color = Color.FromRgb(0x9C, 0x65, 0x00), Bold = true },
        };
        foreach (Cell hit in sheet.FindAll("ESCALATED"))
        {
            hit.Style = flag;
        }
        return wb;
    }

    /// <summary>Regular expressions reach the patterns wildcards cannot.</summary>
    private static Workbook RegexReplace()
    {
        Workbook wb = Tickets();
        Sheet sheet = wb.Sheets[0];

        // Re-number every invoice from INV-1234 to 2026/1234.
        int changed = sheet.ReplaceAll(@"INV-(\d{4})", "2026/$1", new FindOptions { UseRegex = true });
        Note(sheet, "replaceAllRegex(\"INV-(\\\\d{4})\", \"2026/$1\") rewrote "
                + changed + " references.");
        return wb;
    }

    private static void Count(Sheet sheet, int row, string call, int hits)
    {
        sheet[row, 0].Value = call;
        sheet[row, 1].Value = hits;
    }

    private static void Note(Sheet sheet, string text)
    {
        int row = sheet.UsedRange!.LastRow + 2;
        sheet[row, 0].Value = text;
        sheet[row, 0].Style = new CellStyle { Font = new() { Italic = true } };
    }

    /// <summary>
    /// The ticket list, read from a packaged workbook.
    /// </summary>
    /// <remarks>
    /// The used block is read and written straight back, as the Java demo does. The Java
    /// demo says find and replace match only cells written through the API; with this
    /// package they find imported cells too (6 "overdue" hits either way, in .NET and in
    /// Java — see CONVERSION-NOTES/featuredemo-output.md), so the rewrite is kept only to
    /// stay step-for-step with Java. It is value-for-value, so nothing about the sheet
    /// changes.
    /// </remarks>
    private static Workbook Tickets()
    {
        Workbook wb = Workbook.Open(DemoBooks.Read(Content), WorkbookFormat.Xlsx);
        Sheet sheet = wb.Sheets[0];
        CellRange used = sheet.UsedRange!;
        sheet.GetRange(used.FirstRow, used.FirstColumn, used.LastRow, used.LastColumn)
                .SetValues(used.GetValues());
        return wb;
    }
}
