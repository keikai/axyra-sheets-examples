using Axyra.Sheets.IO;
using Axyra.Sheets.Style;
using Axyra.Sheets.Template;

namespace Axyra.Sheets.Examples.Guides;

/// <summary>
/// The complete program from "How to Create Excel Reports Automatically in .NET".
/// It builds the template itself so that it runs without any input files.
/// </summary>
public sealed class GenerateReportExample : IExample
{
    public sealed record Line(string Region, string Product, int Qty, double Amount);

    /// <summary>A 1x1 PNG standing in for assets/logo.png.</summary>
    private static readonly byte[] Logo = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==");

    public string Name => "guides/generate-report";

    public string Description => "Fill an Excel template with Template Markers and render a report";

    public void Run(string outputDirectory)
    {
        string template = CreateTemplate(Path.Combine(outputDirectory, "sales-template.xlsx"));
        string quarter = "Q3-2026";
        List<Line> lines =
        [
            new Line("APAC", "Widget", 12, 51.00),
            new Line("APAC", "Gadget", 3, 27.00),
            new Line("EMEA", "Widget", 8, 34.00),
        ];

        var data = new Dictionary<string, object?>
        {
            ["title"] = quarter,
            ["generated"] = new DateOnly(2026, 9, 9),
            ["lines"] = lines,
            ["logo"] = Logo,
        };

        string xlsx = Path.Combine(outputDirectory, "sales-" + quarter + ".xlsx");
        string pdf = Path.Combine(outputDirectory, "sales-" + quarter + ".pdf");

        using (Workbook wb = Workbook.Open(template))
        {
            Sheet sheet = wb.Sheets["Sales Report"];   // by name, not index

            TemplateMarkers.Process(sheet, data);
            wb.Recalculate();   // refresh template formulas that Process did not touch

            sheet.PageSetup.PrintTitleRows = "$1:$2";   // title + header on every page
            wb.Save(xlsx);

            wb.SavePdf(pdf, new PdfOptions
            {
                Title = "Sales Report " + quarter,
                FitTo = PdfFit.Width,
                Bookmarks = true,
            });

            Console.WriteLine("Total = " + sheet[5, 3].FormattedText);   // 112.00
        }
        Console.WriteLine("Wrote " + xlsx + " and " + pdf);
    }

    /// <summary>The template from the guide: title row, header row, template row, total row.</summary>
    private static string CreateTemplate(string path)
    {
        using (Workbook workbook = new Workbook())
        {
            Sheet sheet = workbook.Sheets.Add("Sales Report");
            sheet["A1:D4"].SetValues(new CellValue[,]
            {
                { "Sales Report — ${title}", CellValue.Empty, "${generated}", "${image:logo}" },
                { "Region", "Product", "Qty", "Amount" },
                { "${lines.region}", "${lines.product}", "${lines.qty}", "${lines.amount}" },
                { "Total", CellValue.Empty, CellValue.Empty, "${lines.amount:sum}" },
            });
            // Formatting on the template row is copied to every inserted row.
            var money = new CellStyle { NumberFormat = "#,##0.00" };
            sheet["D3:D4"].Style = money;
            sheet["A2:D2"].Style = new CellStyle { Font = new Font { Bold = true } };
            workbook.Save(path);
        }
        return path;
    }
}
