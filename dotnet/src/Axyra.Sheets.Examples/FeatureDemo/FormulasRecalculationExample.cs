using Axyra.Sheets.Calc;
using Axyra.Sheets.Formula;
using Axyra.Sheets.Style;

namespace Axyra.Sheets.Examples.FeatureDemo;

/// <summary>Formulas and recalculation.</summary>
/// <remarks>
/// Setting a formula does not evaluate it — <c>Recalculate()</c> does, and
/// <c>RecalculateDirty()</c> does only the part that changed. The engine also evaluates ad hoc,
/// with no cell involved, and supports the modern dynamic-array functions that spill into
/// neighbouring cells.
/// </remarks>
public sealed class FormulasRecalculationExample : IExample
{
    private const string Content = "engineFormulas-invoice-content.xlsx";

    public string Name => "feature-demo/formulas-recalculation";

    public string Description => "Calculate a sheet, evaluate ad hoc, spill dynamic arrays and use manual calc mode";

    public void Run(string outputDirectory)
    {
        using (Workbook wb = Calculate())
        {
            wb.Save(Path.Combine(outputDirectory, "formulas-calculate.xlsx"));
        }
        using (Workbook wb = BeforeAndAfter())
        {
            wb.Save(Path.Combine(outputDirectory, "formulas-before-after.xlsx"));
        }
        using (Workbook wb = AdHoc())
        {
            wb.Save(Path.Combine(outputDirectory, "formulas-ad-hoc.xlsx"));
        }
        using (Workbook wb = DynamicArrays())
        {
            wb.Save(Path.Combine(outputDirectory, "formulas-dynamic-arrays.xlsx"));
        }
        using (Workbook wb = ManualMode())
        {
            wb.Save(Path.Combine(outputDirectory, "formulas-manual-mode.xlsx"));
        }
        Console.WriteLine("Wrote five formula workbooks to " + outputDirectory);
    }

    /// <summary>An ordinary calculated sheet: per-row formulas plus aggregates.</summary>
    private static Workbook Calculate()
    {
        Workbook wb = Invoice();
        Sheet sheet = wb.Sheets[0];

        for (int row = 1; row <= 5; row++)
        {
            sheet[row, 3].Formula = "B" + (row + 1) + "*C" + (row + 1);
        }
        sheet[7, 0].Value = "Subtotal";
        sheet[7, 3].Formula = "SUM(D2:D6)";
        sheet[8, 0].Value = "VAT 20%";
        sheet[8, 3].Formula = "D8*0.2";
        sheet[9, 0].Value = "Total";
        sheet[9, 3].Formula = "D8+D9";
        sheet["A8:D10"].Style = new CellStyle { Font = new Font { Bold = true } };
        sheet["D2:D10"].Style = new CellStyle { NumberFormat = "#,##0.00" };

        wb.Recalculate();
        return wb;
    }

    /// <summary>The same book, showing that a formula is inert until recalculation.</summary>
    private static Workbook BeforeAndAfter()
    {
        Workbook wb = Invoice();
        Sheet sheet = wb.Sheets[0];
        for (int row = 1; row <= 5; row++)
        {
            sheet[row, 3].Formula = "B" + (row + 1) + "*C" + (row + 1);
        }

        sheet[7, 0].Value = "D2 before Recalculate()";
        sheet[7, 3].Value = sheet[1, 3].Value.ToString();

        wb.Recalculate();

        sheet[8, 0].Value = "D2 after Recalculate()";
        sheet[8, 3].Value = sheet[1, 3].Value.ToString();
        sheet.Columns[0].Width = 26;
        sheet.Columns[3].Width = 26;
        return wb;
    }

    /// <summary><c>Evaluate</c> needs no cell — useful for validating user input.</summary>
    private static Workbook AdHoc()
    {
        Workbook wb = new Workbook();
        Sheet sheet = wb.Sheets.Add("Evaluate");
        sheet["A1:A5"].SetNumbers([10, 20, 30, 40, 50]);
        sheet[0, 1].Value = "<- source data";

        string[] expressions =
        [
            "1+2*3",
            "TEXT(0.1234, \"0.0%\")",
            "SUM(A1:A5)",
            "AVERAGE(A1:A5)",
            "XLOOKUP(30, A1:A5, A1:A5)",
            "LET(n, SUM(A1:A5), n/COUNT(A1:A5))",
        ];

        sheet["C1:D1"].SetValues(new CellValue[,] { { "Expression", "Evaluate(...)" } });
        sheet["C1:D1"].Style = new CellStyle { Font = new Font { Bold = true } };
        for (int i = 0; i < expressions.Length; i++)
        {
            sheet[i + 1, 2].Value = expressions[i];
            sheet[i + 1, 3].Value = sheet.Evaluate(expressions[i]).ToString();
        }

        sheet[9, 2].Value = "Functions in the catalogue";
        sheet[9, 3].Value = BuiltInFunctions.Count;
        sheet.Columns[2].Width = 38;
        sheet.Columns[3].Width = 30;
        return wb;
    }

    /// <summary>One formula, many cells: the result spills into the neighbours.</summary>
    private static Workbook DynamicArrays()
    {
        Workbook wb = new Workbook();
        Sheet sheet = wb.Sheets.Add("Spill");

        sheet["A1:B1"].SetValues(new CellValue[,] { { "Region", "Revenue" } });
        string[] regions = ["North", "South", "East", "West", "North", "East"];
        double[] revenue = [1800, 1470, 2415, 1695, 2200, 1310];
        for (int i = 0; i < regions.Length; i++)
        {
            sheet[i + 1, 0].Value = regions[i];
            sheet[i + 1, 1].Value = revenue[i];
        }
        sheet["A1:B1"].Style = new CellStyle { Font = new Font { Bold = true } };

        sheet[0, 3].Value = "SEQUENCE(6)";
        sheet[1, 3].Formula = "SEQUENCE(6)";

        sheet[0, 4].Value = "UNIQUE(A2:A7)";
        sheet[1, 4].Formula = "UNIQUE(A2:A7)";

        sheet[0, 5].Value = "SORT(B2:B7,1,-1)";
        sheet[1, 5].Formula = "SORT(B2:B7,1,-1)";

        sheet[0, 6].Value = "FILTER(A2:A7,B2:B7>1700)";
        sheet[1, 6].Formula = "FILTER(A2:A7,B2:B7>1700)";

        sheet["D1:G1"].Style = new CellStyle { Font = new Font { Bold = true } };
        for (int c = 3; c <= 6; c++)
        {
            sheet.Columns[c].Width = 20;
        }
        wb.Recalculate();
        return wb;
    }

    /// <summary>Manual mode defers evaluation until you ask, like Excel's F9.</summary>
    private static Workbook ManualMode()
    {
        Workbook wb = new Workbook();
        Sheet sheet = wb.Sheets.Add("Manual");
        wb.CalcMode = CalcMode.Manual;

        sheet["A1:A5"].SetNumbers([3, 1, 4, 1, 5]);
        sheet[6, 0].Value = "Calc mode";
        sheet[6, 1].Value = wb.CalcMode.ToString();
        sheet[7, 0].Value = "SUM(A1:A5)";
        sheet[7, 1].Formula = "SUM(A1:A5)";

        sheet[9, 0].Value = "In Manual mode the total stays stale until Recalculate() runs.";
        wb.Recalculate();
        sheet[10, 0].Value = "Recalculate() has now run.";
        sheet.Columns[0].Width = 30;
        return wb;
    }

    /// <summary>The starting sheet: invoice lines read from a packaged workbook.</summary>
    private static Workbook Invoice() => Workbook.Open(DemoBooks.Read(Content), WorkbookFormat.Xlsx);
}
