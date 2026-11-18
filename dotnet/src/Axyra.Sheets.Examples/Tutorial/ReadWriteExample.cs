using Axyra.Sheets.IO;

namespace Axyra.Sheets.Examples.Tutorial;

public sealed class ReadWriteExample : IExample
{
    public string Name => "tutorial/read-write";

    public string Description => "Open, edit, recalculate, and save an existing workbook";

    public void Run(string outputDirectory)
    {
        string source = ExampleSupport.CreateSalesWorkbook(Path.Combine(outputDirectory, "report.xlsx"));
        string edited = Path.Combine(outputDirectory, "report-edited.xlsx");

        using (Workbook workbook = Workbook.Open(source))
        {
            foreach (Sheet each in workbook.Sheets)
            {
                Console.WriteLine("Sheet: " + each.Name);
            }

            Sheet sheet = workbook.Sheets["Q3 Sales"];   // by name, not index
            sheet[2, 2].Value = 1499;
            CellRange q1 = sheet["B2:B4"];
            q1.SetNumbers(new double[,] { { 1200 }, { 980 }, { 1610 } });
            sheet.Columns[0].Width = 24;
            sheet.FreezePanes(1, 0);

            // Edits recalculate their dependents; D2:D4 are already current.
            Console.WriteLine("D3 after edit: " + sheet[2, 3].Value);
            CellValue answer = workbook.Evaluate("SUM('Q3 Sales'!B2:B4)");
            Console.WriteLine("Q1 total: " + answer);

            workbook.Save(edited, new SaveOptions
            {
                FormulaPolicy = FormulaPolicy.Keep,
                SkipComments = false,
                SkipHyperlinks = false,
            });
        }
        Console.WriteLine("Wrote " + edited);
    }
}
