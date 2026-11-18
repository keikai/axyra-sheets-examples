using Axyra.Sheets.Style;

namespace Axyra.Sheets.Examples.DevRef;

public sealed class CellAndRangeExample : IExample
{
    internal sealed record Sale(string Region, double Q1, double Q2);

    public string Name => "dev-ref/cell-and-range";

    public string Description => "Use typed cells, A1 ranges, bulk access, sorting, and record import";

    public void Run(string outputDirectory)
    {
        using Workbook workbook = new Workbook();
        Sheet sheet = workbook.Sheets.Add("Sales");
        var bold = new CellStyle { Font = new Font { Bold = true } };
        sheet.ImportData(0, 0, new List<Sale>
            {
                new Sale("EMEA", 980, 1100),
                new Sale("APAC", 1200, 1450),
            },
            new ImportOptions
            {
                Columns = ["Region", "Q1", "Q2"],
                Headers = true,
                ColumnLabels = ["Region", "Q1", "Q2"],
                HeaderStyle = bold,
            });

        CellRange values = sheet["B2:C3"];
        double[,] numbers = values.GetNumbers();
        values.SetNumbers(numbers);
        sheet["A2:C3"].Sort(new SortKey(0, true));

        Cell date = sheet[4, 0];
        date.DateValue = new DateOnly(2026, 3, 1);
        date.Style = new CellStyle { NumberFormat = "yyyy-mm-dd" };
        sheet[1, 3].Formula = "B2+C2";
        sheet["D2:D2"].AutoFill(sheet["D2:D3"]);
        workbook.Recalculate();

        CellValue calculated = sheet[1, 3].Value;
        Console.WriteLine(Describe(calculated));
        workbook.Save(Path.Combine(outputDirectory, "cell-and-range.xlsx"));
    }

    private static string Describe(CellValue value) => value switch
    {
        CellValue.Number n => "number " + n.Value,
        CellValue.Text t => "text " + t.Value,
        CellValue.Bool b => "bool " + b.Value,
        CellValue.Error e => "error " + e.Display,
        CellValue.Array a => a.RowCount + "x" + a.ColumnCount + " array",
        _ => "blank",
    };
}
