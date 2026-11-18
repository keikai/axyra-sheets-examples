using Axyra.Sheets.Style;

namespace Axyra.Sheets.Examples.FeatureDemo;

/// <summary>Bulk range read and write — the central design choice of the API.</summary>
/// <remarks>
/// <c>SetValues</c> / <c>SetNumbers</c> / <c>GetValues()</c> / <c>GetNumbers()</c> cross into
/// the native engine <b>once per call</b>, not once per cell. A 20x5 block is one crossing, not
/// a hundred; the <c>double</c> path skips boxing entirely. Writing the same block cell by cell
/// is the thing to avoid.
/// </remarks>
public sealed class BulkRangeExample : IExample
{
    private const int Rows = 20;

    public string Name => "feature-demo/bulk-range";

    public string Description => "Write and read whole blocks in one call, including the unboxed numeric path";

    public void Run(string outputDirectory)
    {
        using (Workbook wb = MixedBlock())
        {
            wb.Save(Path.Combine(outputDirectory, "bulk-range-mixed-block.xlsx"));
        }
        using (Workbook wb = NumericFastPath())
        {
            wb.Save(Path.Combine(outputDirectory, "bulk-range-numeric.xlsx"));
        }
        using (Workbook wb = ReadBlockBack())
        {
            wb.Save(Path.Combine(outputDirectory, "bulk-range-read-back.xlsx"));
        }
        Console.WriteLine("Wrote three bulk-range workbooks to " + outputDirectory);
    }

    /// <summary>One <c>CellValue[,]</c>, one crossing, mixed types.</summary>
    private static Workbook MixedBlock()
    {
        Workbook wb = new Workbook();
        Sheet sheet = wb.Sheets.Add("Block");

        CellValue[,] block = new CellValue[Rows + 1, 4];
        block[0, 0] = "Order";
        block[0, 1] = "Units";
        block[0, 2] = "Priority";
        block[0, 3] = "Cleared";
        for (int i = 0; i < Rows; i++)
        {
            block[i + 1, 0] = $"ORD-{1001 + i:D4}";
            block[i + 1, 1] = 20 + (i * 7) % 130;
            block[i + 1, 2] = i % 3 == 0 ? "High" : i % 3 == 1 ? "Normal" : "Low";
            block[i + 1, 3] = i % 4 != 0;
        }
        sheet.GetRange(0, 0, Rows, 3).SetValues(block);
        sheet["A1:D1"].Style = new CellStyle { Font = new Font { Bold = true } };
        sheet.Columns[0].Width = 14;
        sheet.Columns[2].Width = 12;
        sheet.Columns[3].Width = 12;

        Note(sheet, 0, 5, "One SetValues call wrote " + (Rows + 1) * 4 + " cells "
            + "in a single native crossing.");
        return wb;
    }

    /// <summary>A flat <c>double</c> span, no boxing — the fastest way to fill numbers.</summary>
    private static Workbook NumericFastPath()
    {
        Workbook wb = new Workbook();
        Sheet sheet = wb.Sheets.Add("Numbers");

        sheet["A1:C1"].SetValues(new CellValue[,] { { "Month", "Plan", "Actual" } });
        sheet["A1:C1"].Style = new CellStyle { Font = new Font { Bold = true } };

        string[] months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
            "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
        for (int i = 0; i < months.Length; i++)
        {
            sheet[i + 1, 0].Value = months[i];
        }

        // SetNumbers(ReadOnlySpan<double>) fills row-major across the whole rectangle.
        double[] figures = new double[months.Length * 2];
        for (int i = 0; i < months.Length; i++)
        {
            figures[i * 2] = 10_000 + i * 450;
            figures[i * 2 + 1] = 9_400 + ((i * 1373) % 3100);
        }
        sheet.GetRange(1, 1, months.Length, 2).SetNumbers(figures);
        sheet.GetRange(1, 1, months.Length, 2).Style = new CellStyle { NumberFormat = "#,##0" };

        Note(sheet, 0, 4, "SetNumbers(ReadOnlySpan<double>) writes the rectangle row-major, unboxed.");
        return wb;
    }

    /// <summary>Reading is symmetrical: one call returns the whole rectangle.</summary>
    private static Workbook ReadBlockBack()
    {
        Workbook wb = NumericFastPath();
        Sheet sheet = wb.Sheets[0];

        CellRange block = sheet.GetRange(1, 1, 12, 2);
        double[,] read = block.GetNumbers();

        double planTotal = 0;
        double actualTotal = 0;
        for (int r = 0; r < read.GetLength(0); r++)
        {
            planTotal += read[r, 0];
            actualTotal += read[r, 1];
        }

        sheet[14, 0].Value = "Read back";
        sheet["A15"].Style = new CellStyle { Font = new Font { Bold = true } };
        sheet[15, 0].Value = "values read";
        sheet[15, 1].Value = read.Length;
        sheet[16, 0].Value = "plan total";
        sheet[16, 1].Value = planTotal;
        sheet[17, 0].Value = "actual total";
        sheet[17, 1].Value = actualTotal;
        sheet[18, 0].Value = "variance";
        sheet[18, 1].Value = actualTotal - planTotal;
        sheet.GetRange(15, 1, 18, 1).Style = new CellStyle { NumberFormat = "#,##0" };

        Note(sheet, 20, 0, "Totals computed in C# from one GetNumbers() call — "
            + "no per-cell reads, no formulas involved.");
        return wb;
    }

    private static void Note(Sheet sheet, int row, int col, string text)
    {
        sheet[row, col].Value = text;
        sheet.GetRange(row, col, row, col).Style = new CellStyle { Font = new Font { Italic = true } };
    }
}
