using Axyra.Sheets.IO;

namespace Axyra.Sheets.Examples.DevRef;

public sealed class StreamingExample : IExample
{
    public string Name => "dev-ref/streaming";

    public string Description => "Write rows forward-only and use the columnar workbook backend";

    public void Run(string outputDirectory)
    {
        string streamed = Path.Combine(outputDirectory, "streamed.xlsx");
        using (StreamWorkbook output = new StreamWorkbook(streamed))
        {
            int header = output.AddStyle(new StreamStyle { Bold = true });
            int money = output.AddStyle(new StreamStyle { NumberFormat = "#,##0.00" });
            int date = output.AddStyle(new StreamStyle { NumberFormat = "yyyy-mm-dd" });
            output.StartSheet("Data");
            output.StartRow(0).AddText("Date", header).AddText("Region", header)
                .AddText("Amount", header).Commit();
            for (int i = 0; i < 10; i++)
            {
                output.StartRow(i + 1).AddDate(new DateOnly(2026, 9, 1).AddDays(i), date)
                    .AddText(i % 2 == 0 ? "APAC" : "EMEA")
                    .AddNumber(1000 + i * 25, money).Commit();
            }
            output.Finish();
        }

        using (Workbook workbook = new Workbook(WorkbookStorage.Columnar))
        {
            Sheet sheet = workbook.Sheets.Add("Columnar");
            sheet["A1:A1000"].SetNumbers(new double[1000]);
            workbook.Save(Path.Combine(outputDirectory, "columnar.xlsx"));
        }
    }
}
