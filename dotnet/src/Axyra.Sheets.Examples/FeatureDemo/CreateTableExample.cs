using Axyra.Sheets.Style;

namespace Axyra.Sheets.Examples.FeatureDemo;

/// <summary>
/// Create a table — Excel's "Format as Table", from .NET.
/// </summary>
/// <remarks>
/// A table is more than banded styling: its columns get names, and those names become
/// <b>structured references</b>, so <c>SUM(Orders[Amount])</c> keeps working when rows
/// are added. Header and totals rows are toggled on the table.
/// </remarks>
public sealed class CreateTableExample : IExample
{
    private const string Content = "createTable-orders-content.xlsx";

    public string Name => "feature-demo/create-table";

    public string Description => "Create a table, use structured references and toggle header and totals rows";

    public void Run(string outputDirectory)
    {
        using (Workbook wb = Create())
        {
            wb.Save(Path.Combine(outputDirectory, "create-table-named.xlsx"));
        }
        using (Workbook wb = Inferred())
        {
            wb.Save(Path.Combine(outputDirectory, "create-table-inferred.xlsx"));
        }
        using (Workbook wb = Structured())
        {
            wb.Save(Path.Combine(outputDirectory, "create-table-structured.xlsx"));
        }
        using (Workbook wb = HeaderAndTotals())
        {
            wb.Save(Path.Combine(outputDirectory, "create-table-totals.xlsx"));
        }
        Console.WriteLine("Wrote four table workbooks to " + outputDirectory);
    }

    /// <summary>Explicit column names, given at creation.</summary>
    private static Workbook Create()
    {
        Workbook wb = Orders();
        Sheet sheet = wb.Sheets[0];
        sheet.Tables.Add("Orders", sheet.GetRange(0, 0, 6, 3), "Item", "Qty", "Price", "Amount");
        Note(sheet, 8, "createTable(\"Orders\", \"Item\", \"Qty\", \"Price\", \"Amount\")");
        wb.Recalculate();
        return wb;
    }

    /// <summary>With no names given, the header row supplies them.</summary>
    private static Workbook Inferred()
    {
        Workbook wb = Orders();
        Sheet sheet = wb.Sheets[0];
        sheet.Tables.Add("Orders", sheet.GetRange(0, 0, 6, 3));

        // The names came from row 1, so structured references work immediately.
        Note(sheet, 8, "createTable(\"Orders\") took the column names from row 1:");
        sheet[9, 0].Value = "SUM(Orders[Amount])";
        sheet[9, 3].Formula = "SUM(Orders[Amount])";
        sheet[9, 3].Style = new CellStyle { Font = new() { Bold = true }, NumberFormat = "#,##0.00" };
        wb.Recalculate();
        return wb;
    }

    /// <summary>The point of naming the columns: formulas that survive edits.</summary>
    private static Workbook Structured()
    {
        Workbook wb = Orders();
        Sheet sheet = wb.Sheets[0];
        sheet.Tables.Add("Orders", sheet.GetRange(0, 0, 6, 3), "Item", "Qty", "Price", "Amount");

        sheet[8, 0].Value = "SUM(Orders[Amount])";
        sheet[8, 3].Formula = "SUM(Orders[Amount])";
        sheet[9, 0].Value = "AVERAGE(Orders[Price])";
        sheet[9, 3].Formula = "AVERAGE(Orders[Price])";
        sheet[10, 0].Value = "COUNTA(Orders[Item])";
        sheet[10, 3].Formula = "COUNTA(Orders[Item])";
        sheet.GetRange(8, 0, 10, 0).Style = new CellStyle { Font = new() { Bold = true } };
        sheet.GetRange(8, 3, 9, 3).Style = new CellStyle { NumberFormat = "#,##0.00" };

        Note(sheet, 12, "These reference the table by name, not by A1 address — "
                + "add a row and they still cover it.");
        wb.Recalculate();
        return wb;
    }

    /// <summary>Both rows are switches on the table, which is looked up by name.</summary>
    private static Workbook HeaderAndTotals()
    {
        Workbook wb = Orders();
        Sheet sheet = wb.Sheets[0];
        sheet.Tables.Add("Orders", sheet.GetRange(0, 0, 6, 3), "Item", "Qty", "Price", "Amount");

        sheet.Tables["Orders"].ShowHeaderRow = true;
        sheet.Tables["Orders"].ShowTotalsRow = true;

        Note(sheet, 9, "setTableTotalsRow(\"Orders\", true) added the totals row below the data.");
        wb.Recalculate();
        return wb;
    }

    private static void Note(Sheet sheet, int row, string text)
    {
        sheet[row, 0].Value = text;
        sheet[row, 0].Style = new CellStyle { Font = new() { Italic = true } };
    }

    /// <summary>The starting sheet: order lines read from a packaged workbook.</summary>
    private static Workbook Orders() => Workbook.Open(DemoBooks.Read(Content), WorkbookFormat.Xlsx);
}
