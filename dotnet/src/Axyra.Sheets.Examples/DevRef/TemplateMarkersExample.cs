using Axyra.Sheets.Template;

namespace Axyra.Sheets.Examples.DevRef;

public sealed class TemplateMarkersExample : IExample
{
    internal sealed record Line(string Item, int Qty, double Amount);

    public string Name => "dev-ref/template-markers";

    public string Description => "Fill scalar, list, and aggregate Template Markers";

    public void Run(string outputDirectory)
    {
        using Workbook workbook = new Workbook();
        Sheet sheet = workbook.Sheets.Add("Invoice");
        sheet["A1:C4"].SetValues(new CellValue[,]
        {
            { "Invoice for ${customer.name}", CellValue.Empty, CellValue.Empty },
            { "Item", "Qty", "Amount" },
            { "${lines.item}", "${lines.qty}", "${lines.amount}" },
            { "Total", CellValue.Empty, "${lines.amount:sum}" },
        });
        var data = new Dictionary<string, object?>
        {
            ["customer"] = new Dictionary<string, object?> { ["name"] = "Acme Corp" },
            ["lines"] = new List<Line> { new Line("Widget", 12, 51), new Line("Gadget", 3, 27) },
        };
        TemplateMarkers.Process(sheet, data);
        workbook.RecalculateDirty();
        workbook.Save(Path.Combine(outputDirectory, "template-markers.xlsx"));
    }
}
