using Axyra.Sheets.Charts;
using Axyra.Sheets.Drawing;
using Axyra.Sheets.Style;

namespace Axyra.Sheets.Examples.DevRef;

public sealed class ChartsExample : IExample
{
    public string Name => "dev-ref/charts";

    public string Description => "Build an embedded chart and a chart sheet";

    public void Run(string outputDirectory)
    {
        using Workbook workbook = new Workbook();
        Sheet sheet = workbook.Sheets.Add("Sales");
        sheet["A1:C4"].SetValues(new CellValue[,]
        {
            { "Region", "Q1", "Q2" },
            { "APAC", 1200, 1450 },
            { "EMEA", 980, 1100 },
            { "Americas", 1610, 1720 },
        });
        Chart chart = sheet.Charts.Add(ChartType.Column, new DrawingAnchor(1, 5, 16, 12));
        Configure(chart);

        // A chart sheet is added empty with its type; the same definition is applied to it.
        ChartSheet chartSheet = workbook.ChartSheets.Add("Revenue", ChartType.Column);
        Chart view = chartSheet.Chart!;
        Configure(view);
        Console.WriteLine("Chart sheet: " + chartSheet.Name + "; title=" + view.Title);
        workbook.Save(Path.Combine(outputDirectory, "charts.xlsx"));
    }

    private static void Configure(Chart chart)
    {
        chart.Title = "Revenue by Region";
        chart.Series.Add("Q1", "Sales!$B$2:$B$4", "Sales!$A$2:$A$4");
        chart.Series.Add("Q2", "Sales!$C$2:$C$4", "Sales!$A$2:$A$4");
        chart.Legend.Position = LegendPosition.Bottom;
        chart.ShowDataLabels = true;
        chart.ChartArea.FillColor = Color.Parse("#FFFFFF");
    }
}
