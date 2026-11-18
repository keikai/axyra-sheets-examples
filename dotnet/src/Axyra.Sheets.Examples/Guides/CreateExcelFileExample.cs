using Axyra.Sheets.Style;

namespace Axyra.Sheets.Examples.Guides;

/// <summary>The complete program from "How to Create an Excel File in .NET".</summary>
public sealed class CreateExcelFileExample : IExample
{
    public string Name => "guides/create-excel-file";

    public string Description => "Create a styled sales workbook with formulas and frozen headers";

    public void Run(string outputDirectory)
    {
        string output = Path.Combine(outputDirectory, "q3-sales.xlsx");
        using (Workbook wb = new Workbook())
        {
            Sheet sheet = wb.Sheets.Add("Q3 Sales");

            // Header and data, one call each.
            sheet["A1:D1"].SetValues(new CellValue[,]
            {
                { "Region", "Q1", "Q2", "Total" },
            });
            sheet["A2:C4"].SetValues(new CellValue[,]
            {
                { "APAC",     1200, 1450 },
                { "EMEA",     980,  1100 },
                { "Americas", 1610, 1720 },
            });

            // Per-row totals, then a grand-total row. References shift per cell.
            sheet["D2:D4"].SetFormula("B2+C2");
            sheet[4, 0].Value = "Total";
            sheet["B5:D5"].SetFormula("SUM(B2:B4)");

            // Styling.
            var header = new CellStyle
            {
                Font = new Font { Bold = true, Color = Color.Parse("#FFFFFF") },
                Fill = Color.Parse("#2F5597")!,   // Parse returns null for text it cannot read
                HorizontalAlignment = HorizontalAlignment.Center,
                Borders = Borders.Around(BorderStyle.Thin, Color.Parse("#1F3864")!),
            };
            var money = new CellStyle { NumberFormat = "#,##0.00" };
            var totalRow = new CellStyle
            {
                Font = new Font { Bold = true },
                NumberFormat = "#,##0.00",
            };

            sheet["A1:D1"].Style = header;
            sheet["B2:D4"].Style = money;
            sheet["A5:D5"].Style = totalRow;

            // Layout.
            sheet.Columns[0].Width = 14.0;
            for (int col = 1; col <= 3; col++)
            {
                sheet.Columns[col].AutoFit();
            }
            sheet.FreezePanes(1, 0);

            wb.Save(output);
            Console.WriteLine("D5 = " + sheet[4, 3].Value);   // 8060
        }
        Console.WriteLine("Wrote " + output);
    }
}
