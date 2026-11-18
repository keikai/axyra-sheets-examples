namespace Axyra.Sheets.Examples.Tutorial;

public sealed class QuickStartExample : IExample
{
    public string Name => "tutorial/quick-start";

    public string Description => "Create a workbook, calculate a formula, and save hello.xlsx";

    public void Run(string outputDirectory)
    {
        string output = Path.Combine(outputDirectory, "hello.xlsx");
        using (Workbook wb = new Workbook())
        {
            Sheet sheet = wb.Sheets.Add("Sheet1");

            sheet[0, 0].Value = "Item";
            sheet[0, 1].Value = "Qty";
            sheet[0, 2].Value = "Price";

            sheet[1, 0].Value = "Widget";
            sheet[1, 1].Value = 12;
            sheet[1, 2].Value = 4.25;

            sheet[1, 3].Formula = "B2*C2";   // calculated right away

            wb.Save(output);

            // Step 3: read a value back.
            CellValue total = sheet[1, 3].Value;
            if (total is CellValue.Number number)
            {
                Console.WriteLine(number.Value);              // 51
            }
            Console.WriteLine(sheet[1, 3].FormattedText);     // 51

            // Step 4: A1 notation.
            CellRange header = sheet["A1:D1"];   // same as sheet.GetRange(0, 0, 0, 3)
            Console.WriteLine("Header columns: " + header.ColumnCount);
        }
        Console.WriteLine("Wrote " + output);
    }
}
