using System.Globalization;
using Axyra.Sheets.Style;

namespace Axyra.Sheets.Examples.FeatureDemo;

/// <summary>Write and read a single cell.</summary>
/// <remarks>
/// <see cref="CellValue"/> is a closed hierarchy — <c>Number</c>, <c>Text</c>, <c>Bool</c>,
/// <c>Blank</c> and friends — so the kind that went in is the kind that comes back. Dates go in
/// as <see cref="DateOnly"/> / <see cref="DateTime"/> values over the Excel serial and come back
/// the same way.
/// </remarks>
public sealed class CellReadWriteExample : IExample
{
    private const string Content = "cellReadWrite-header-content.xlsx";

    public string Name => "feature-demo/cell-read-write";

    public string Description => "Write every CellValue kind, read it back, and round-trip dates and times";

    public void Run(string outputDirectory)
    {
        using (Workbook wb = WriteTypes())
        {
            wb.Save(Path.Combine(outputDirectory, "cell-read-write-types.xlsx"));
        }
        using (Workbook wb = ReadBack())
        {
            wb.Save(Path.Combine(outputDirectory, "cell-read-write-read-back.xlsx"));
        }
        using (Workbook wb = DatesAndTimes())
        {
            wb.Save(Path.Combine(outputDirectory, "cell-read-write-dates.xlsx"));
        }
        Console.WriteLine("Wrote three cell read/write workbooks to " + outputDirectory);
    }

    /// <summary>Every <see cref="CellValue"/> kind, one per row.</summary>
    private static Workbook WriteTypes()
    {
        Workbook wb = Header("Written");
        Sheet sheet = wb.Sheets[0];

        sheet[1, 0].Value = CellValue.Of("text");
        sheet[1, 1].Value = CellValue.Of("Widget");

        sheet[2, 0].Value = CellValue.Of("number");
        sheet[2, 1].Value = CellValue.Of(1234.5);

        sheet[3, 0].Value = CellValue.Of("boolean");
        sheet[3, 1].Value = CellValue.Of(true);

        sheet[4, 0].Value = CellValue.Of("blank");
        sheet[4, 1].Value = CellValue.Empty;

        sheet[5, 0].Value = CellValue.Of("formula");
        sheet[5, 1].Formula = "B3*2";

        // The shorthand for the common case: a bare number converts implicitly, no wrapper.
        sheet[6, 0].Value = "implicit";
        sheet[6, 1].Value = 42;

        wb.Recalculate();
        return wb;
    }

    /// <summary>What comes back out, which variant carried it, and how it displays.</summary>
    private static Workbook ReadBack()
    {
        Workbook wb = WriteTypes();
        Sheet sheet = wb.Sheets[0];
        sheet["C1:D1"].SetValues(new CellValue[,] { { "CellValue variant", "Displayed as" } });
        sheet["C1:D1"].Style = new CellStyle { Font = new Font { Bold = true } };

        for (int row = 1; row <= 6; row++)
        {
            CellValue value = sheet[row, 1].Value;
            sheet[row, 2].Value = value.GetType().Name;
            sheet[row, 3].Value = sheet[row, 1].FormattedText;
        }
        sheet.Columns[2].Width = 22;
        sheet.Columns[3].Width = 22;
        sheet[8, 0].Value = "Every value goes in and comes back out as the same kind.";
        return wb;
    }

    /// <summary><c>DateValue</c> / <c>DateTimeValue</c> write the Excel serial; the format makes it legible.</summary>
    private static Workbook DatesAndTimes()
    {
        Workbook wb = Header("Dates");
        Sheet sheet = wb.Sheets[0];

        sheet[1, 0].Value = "DateOnly";
        sheet[1, 1].DateValue = new DateOnly(2026, 8, 24);
        sheet["B2"].Style = new CellStyle { NumberFormat = "yyyy-mm-dd" };

        sheet[2, 0].Value = "DateTime";
        sheet[2, 1].DateTimeValue = new DateTime(2026, 8, 24, 14, 30, 0);
        sheet["B3"].Style = new CellStyle { NumberFormat = "yyyy-mm-dd hh:mm" };

        sheet[3, 0].Value = "The raw serial";
        sheet[3, 1].Value = sheet[1, 1].Value;

        sheet[5, 0].Value = "DateValue ->";
        sheet[5, 1].Value = sheet[1, 1].DateValue?.ToString("O", CultureInfo.InvariantCulture) ?? "null";
        sheet[6, 0].Value = "DateTimeValue ->";
        sheet[6, 1].Value = sheet[2, 1].DateTimeValue?.ToString("s", CultureInfo.InvariantCulture) ?? "null";

        sheet.Columns[1].Width = 22;
        return wb;
    }

    /// <summary>The starting sheet: a header row read from a packaged workbook, renamed for the sample.</summary>
    private static Workbook Header(string name)
    {
        Workbook wb = Workbook.Open(DemoBooks.Read(Content), WorkbookFormat.Xlsx);
        wb.Sheets[0].Name = name;
        return wb;
    }
}
