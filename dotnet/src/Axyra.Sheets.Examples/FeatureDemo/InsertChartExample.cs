using Axyra.Sheets.Charts;
using Axyra.Sheets.Drawing;
using Axyra.Sheets.Style;

namespace Axyra.Sheets.Examples.FeatureDemo;

/// <summary>
/// Insert a chart.
/// </summary>
/// <remarks>
/// A chart is added with its type and the anchor rectangle it occupies on the sheet, then
/// given its title, series, legend and grouping. Series are bound to ranges by A1
/// address, so the chart redraws whenever those cells change.
/// </remarks>
public sealed class InsertChartExample : IExample
{
    private const string Content = "insertChart-data-content.xlsx";

    public string Name => "feature-demo/insert-chart";

    public string Description => "Insert column, line, area and pie charts bound to sheet ranges";

    public void Run(string outputDirectory)
    {
        using (Workbook wb = Column())
        {
            wb.Save(Path.Combine(outputDirectory, "insert-chart-column.xlsx"));
        }
        using (Workbook wb = Line())
        {
            wb.Save(Path.Combine(outputDirectory, "insert-chart-line.xlsx"));
        }
        using (Workbook wb = Area())
        {
            wb.Save(Path.Combine(outputDirectory, "insert-chart-area.xlsx"));
        }
        using (Workbook wb = Pie())
        {
            wb.Save(Path.Combine(outputDirectory, "insert-chart-pie.xlsx"));
        }
        using (Workbook wb = TwoCharts())
        {
            wb.Save(Path.Combine(outputDirectory, "insert-chart-two-charts.xlsx"));
        }
        Console.WriteLine("Wrote five chart workbooks to " + outputDirectory);
    }

    // The anchors are corner cells, as in the Java builder's anchor(fromRow, fromCol, toRow, toCol):
    // the chart ends at the top-left of the "to" cell.

    private static Workbook Column()
    {
        Workbook wb = Data();
        Sheet sheet = wb.Sheets[0];
        Chart chart = sheet.Charts.Add(ChartType.Column, new DrawingAnchor(1, 4, 18, 12));
        chart.Title = "Plan vs actual";
        chart.Series.Add("Plan", "Data!B2:B7", "Data!A2:A7");
        chart.Series.Add("Actual", "Data!C2:C7", "Data!A2:A7").FillColor = Color.FromRgb(0x1F, 0x4E, 0x79);
        chart.Legend.Position = LegendPosition.Bottom;
        chart.Grouping = ChartGrouping.Clustered;
        chart.GapWidth = 60;
        return wb;
    }

    private static Workbook Line()
    {
        Workbook wb = Data();
        Sheet sheet = wb.Sheets[0];
        Chart chart = sheet.Charts.Add(ChartType.Line, new DrawingAnchor(1, 4, 18, 12));
        chart.Title = "Trend";
        chart.Series.Add("Plan", "Data!B2:B7", "Data!A2:A7");
        chart.Series.Add("Actual", "Data!C2:C7", "Data!A2:A7");
        chart.Legend.Position = LegendPosition.Right;
        return wb;
    }

    private static Workbook Area()
    {
        Workbook wb = Data();
        Sheet sheet = wb.Sheets[0];
        Chart chart = sheet.Charts.Add(ChartType.Area, new DrawingAnchor(1, 4, 18, 12));
        chart.Title = "Cumulative";
        chart.Series.Add("Plan", "Data!B2:B7", "Data!A2:A7");
        chart.Series.Add("Actual", "Data!C2:C7", "Data!A2:A7");
        chart.Grouping = ChartGrouping.Stacked;
        chart.Legend.Position = LegendPosition.Bottom;
        return wb;
    }

    private static Workbook Pie()
    {
        Workbook wb = Data();
        Sheet sheet = wb.Sheets[0];
        Chart chart = sheet.Charts.Add(ChartType.Pie, new DrawingAnchor(1, 4, 18, 12));
        chart.Title = "Actual by month";
        chart.Series.Add("Actual", "Data!C2:C7", "Data!A2:A7");
        chart.Legend.Position = LegendPosition.Right;
        return wb;
    }

    /// <summary>Charts are anchored to a cell rectangle, so several fit on one sheet.</summary>
    private static Workbook TwoCharts()
    {
        Workbook wb = Data();
        Sheet sheet = wb.Sheets[0];

        Chart column = sheet.Charts.Add(ChartType.Column, new DrawingAnchor(1, 4, 15, 10));
        column.Title = "Plan vs actual";
        column.Series.Add("Plan", "Data!B2:B7", "Data!A2:A7");
        column.Series.Add("Actual", "Data!C2:C7", "Data!A2:A7");
        column.Legend.Position = LegendPosition.Bottom;

        Chart pie = sheet.Charts.Add(ChartType.Pie, new DrawingAnchor(16, 4, 30, 10));
        pie.Title = "Share of actual";
        pie.Series.Add("Actual", "Data!C2:C7", "Data!A2:A7");
        pie.Legend.Position = LegendPosition.Right;

        sheet[9, 0].Value = "anchor(firstRow, firstCol, lastRow, lastCol) places each one.";
        return wb;
    }

    /// <summary>The starting sheet: a "Data" sheet of monthly plan and actual figures.</summary>
    private static Workbook Data() => Workbook.Open(DemoBooks.Read(Content), WorkbookFormat.Xlsx);
}
