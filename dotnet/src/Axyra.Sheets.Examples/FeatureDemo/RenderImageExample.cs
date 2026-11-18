using Axyra.Sheets.Charts;
using Axyra.Sheets.Drawing;
using Axyra.Sheets.IO;
using Axyra.Sheets.Style;

namespace Axyra.Sheets.Examples.FeatureDemo;

/// <summary>
/// Render a sheet to an image.
/// </summary>
/// <remarks>
/// <see cref="Sheet.RenderImage(ImageFormat, ImageOptions)"/> rasterises a sheet
/// server-side — PNG, JPEG, GIF, BMP, TIFF or SVG — with no browser and no viewer
/// component involved. That is what makes it useful for thumbnails, email bodies and
/// report previews.
/// </remarks>
public sealed class RenderImageExample : IExample
{
    public string Name => "feature-demo/render-image";

    public string Description => "Render a sheet to PNG at different scales, with or without gridlines, and to SVG";

    public void Run(string outputDirectory)
    {
        using (Workbook wb = Dashboard())
        {
            wb.Save(Path.Combine(outputDirectory, "render-image-dashboard.xlsx"));
        }
        Write(Path.Combine(outputDirectory, "sheet.png"), Render(ImageFormat.Png, 1.0, true));
        Write(Path.Combine(outputDirectory, "sheet-2x.png"), Render(ImageFormat.Png, 2.0, true));
        Write(Path.Combine(outputDirectory, "sheet-no-gridlines.png"), Render(ImageFormat.Png, 1.5, false));
        Write(Path.Combine(outputDirectory, "sheet.svg"), RenderSvg());
    }

    private static void Write(string file, byte[] bytes)
    {
        File.WriteAllBytes(file, bytes);
        Console.WriteLine("Wrote " + Path.GetFileName(file) + " (" + bytes.Length + " bytes)");
    }

    /// <summary>Rasterise the first sheet at the given scale.</summary>
    private static byte[] Render(ImageFormat format, double scale, bool gridlines)
    {
        using Workbook wb = Dashboard();
        return wb.Sheets[0].RenderImage(format, new ImageOptions
        {
            Scale = scale,
            Gridlines = gridlines,
        });
    }

    /// <summary>SVG stays vector.</summary>
    private static byte[] RenderSvg()
    {
        using Workbook wb = Dashboard();
        return wb.Sheets[0].RenderImage(ImageFormat.Svg, new ImageOptions());
    }

    /// <summary>Something worth looking at: figures, a total and a chart.</summary>
    private static Workbook Dashboard()
    {
        Workbook wb = new Workbook();
        Sheet sheet = wb.Sheets.Add("Dashboard");

        sheet["A1:C1"].SetValues(new CellValue[,] { { "Month", "Plan", "Actual" } });
        sheet["A1:C1"].Style = new CellStyle
        {
            Font = new() { Bold = true, Color = Color.White },
            Fill = Color.FromRgb(0x1F, 0x4E, 0x79),
        };

        string[] months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun"];
        double[] plan = [12000, 12500, 13000, 13500, 14000, 14500];
        double[] actual = [11400, 13200, 12750, 15100, 13800, 16250];
        for (int i = 0; i < months.Length; i++)
        {
            sheet[i + 1, 0].Value = months[i];
            sheet[i + 1, 1].Value = plan[i];
            sheet[i + 1, 2].Value = actual[i];
        }
        sheet.GetRange(1, 1, months.Length, 2).Style = new CellStyle { NumberFormat = "#,##0" };

        sheet[8, 0].Value = "Total";
        sheet[8, 1].Formula = "SUM(B2:B7)";
        sheet[8, 2].Formula = "SUM(C2:C7)";
        sheet["A9:C9"].Style = new CellStyle { Font = new() { Bold = true }, NumberFormat = "#,##0" };

        // Corner cells, as Java's anchor(1, 4, 18, 11): the chart ends at the top-left of L19.
        Chart chart = sheet.Charts.Add(ChartType.Column, new DrawingAnchor(1, 4, 18, 11));
        chart.Title = "Plan vs actual";
        chart.Series.Add("Plan", "Dashboard!B2:B7", "Dashboard!A2:A7");
        chart.Series.Add("Actual", "Dashboard!C2:C7", "Dashboard!A2:A7");
        chart.Legend.Position = LegendPosition.Bottom;

        sheet.Columns[0].Width = 12;
        sheet.Columns[1].Width = 12;
        sheet.Columns[2].Width = 12;

        // An image is one printed page, so lay the page out around the content:
        // the table and the chart beside it, landscape, scaled onto a single page.
        sheet.PageSetup.PrintArea = "A1:K19";
        sheet.PageSetup.Orientation = PageOrientation.Landscape;
        sheet.PageSetup.FitToPages(1, 1);
        wb.Recalculate();
        return wb;
    }
}
