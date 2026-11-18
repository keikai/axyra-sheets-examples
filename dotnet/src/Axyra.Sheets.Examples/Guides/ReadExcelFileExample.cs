using System.Text;

namespace Axyra.Sheets.Examples.Guides;

/// <summary>The complete program from "How to Read an Excel File in .NET".</summary>
public sealed class ReadExcelFileExample : IExample
{
    public string Name => "guides/read-excel-file";

    public string Description => "Read each worksheet's used range and print it as TSV";

    public void Run(string outputDirectory)
    {
        string source = ExampleSupport.CreateSalesWorkbook(Path.Combine(outputDirectory, "read-source.xlsx"));
        using Workbook wb = Workbook.Open(source);
        foreach (Sheet sheet in wb.Sheets)
        {
            if (sheet.SheetType != SheetType.Worksheet)
            {
                continue;
            }

            CellRange? used = sheet.UsedRange;
            if (used == null)
            {
                Console.WriteLine("== " + sheet.Name + " (empty)");
                continue;
            }

            Console.WriteLine($"== {sheet.Name}  rows {used.FirstRow}-{used.LastRow}, "
                + $"cols {used.FirstColumn}-{used.LastColumn}");

            CellValue[,] values = used.GetValues();
            for (int r = 0; r < values.GetLength(0); r++)
            {
                var line = new StringBuilder();
                for (int c = 0; c < values.GetLength(1); c++)
                {
                    if (line.Length > 0)
                    {
                        line.Append('\t');
                    }
                    line.Append(Text(values[r, c]));
                }
                Console.WriteLine(line);
            }
        }
    }

    private static string Text(CellValue v) => v switch
    {
        CellValue.Number n => n.Value.ToString(),
        CellValue.Text t => t.Value,
        CellValue.Bool b => b.Value.ToString(),
        CellValue.Error e => e.Display,
        CellValue.Array => "[array]",
        _ => "",
    };
}
