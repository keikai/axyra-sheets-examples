using Axyra.Sheets.Pivots;

namespace Axyra.Sheets.Examples.DevRef;

public sealed class PivotTableExample : IExample
{
    public string Name => "dev-ref/pivot-table";

    public string Description => "Create, configure, filter, and refresh a pivot table";

    public void Run(string outputDirectory)
    {
        using Workbook workbook = new Workbook();
        Sheet data = workbook.Sheets.Add("Data");
        data["A1:E5"].SetValues(new CellValue[,]
        {
            { "Region", "Product", "Quarter", "Channel", "Amount" },
            { "APAC", "Widget", "Q1", "Direct", 1200 },
            { "EMEA", "Widget", "Q1", "Partner", 980 },
            { "APAC", "Gadget", "Q2", "Direct", 1450 },
            { "Americas", "Gadget", "Q2", "Partner", 1720 },
        });
        Sheet summary = workbook.Sheets.Add("Summary");
        PivotTable pivot = summary.PivotTables.Add("SalesPivot", "A1", "Data!A1:E5");
        pivot.RowFields.Add("Region");
        pivot.ColumnFields.Add("Quarter");
        pivot.DataFields.Add("Amount", AggregateFunction.Sum);
        pivot.Fields["Region"].VisibleItems = ["APAC", "EMEA", "Americas"];
        pivot.DataFields[0].ShowDataAs = new PivotShowDataAs(PivotShowDataAsType.PercentOfTotal);
        pivot.Style = new PivotStyle("PivotStyleMedium9") { ShowRowStripes = true };
        int written = pivot.Refresh();
        Console.WriteLine("Pivot cells written: " + written);
        workbook.Save(Path.Combine(outputDirectory, "pivot-table.xlsx"));
    }
}
