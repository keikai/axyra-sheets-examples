namespace Axyra.Sheets.Examples;

public static class ExampleSupport
{
    public static string PrepareOutputDirectory(string directory)
    {
        string full = Path.GetFullPath(directory);
        Directory.CreateDirectory(full);
        return full;
    }

    public static string CreateSalesWorkbook(string output)
    {
        using Workbook workbook = new Workbook();
        Sheet sheet = workbook.Sheets.Add("Q3 Sales");
        sheet["A1:D4"].SetValues(new CellValue[,]
        {
            { "Region", "Q1", "Q2", "Total" },
            { "APAC", 1200, 1450, CellValue.Empty },
            { "EMEA", 980, 1100, CellValue.Empty },
            { "Americas", 1610, 1720, CellValue.Empty },
        });
        sheet["D2:D4"].SetFormula("B2+C2");   // D3 = B3+C3, D4 = B4+C4
        workbook.Save(output);
        return output;
    }
}
