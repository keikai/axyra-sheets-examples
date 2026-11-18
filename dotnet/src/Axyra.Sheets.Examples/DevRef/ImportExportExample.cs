using Axyra.Sheets.IO;

namespace Axyra.Sheets.Examples.DevRef;

public sealed class ImportExportExample : IExample
{
    public string Name => "dev-ref/import-export";

    public string Description => "Open from paths and bytes, then save XLSX, CSV, and JSON";

    public void Run(string outputDirectory)
    {
        string source = ExampleSupport.CreateSalesWorkbook(Path.Combine(outputDirectory, "import-source.xlsx"));
        byte[] bytes = File.ReadAllBytes(source);
        using Workbook workbook = Workbook.Open(bytes, WorkbookFormat.Xlsx);
        workbook.Recalculate();
        workbook.Save(Path.Combine(outputDirectory, "values-only.xlsx"), new SaveOptions
        {
            FormulaPolicy = FormulaPolicy.ValuesOnly,
        });
        workbook.Save(Path.Combine(outputDirectory, "sales.csv"));
        File.WriteAllBytes(Path.Combine(outputDirectory, "sales.json"), workbook.ToByteArray(WorkbookFormat.Json));
    }
}
