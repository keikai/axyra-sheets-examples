using Axyra.Sheets.IO;

namespace Axyra.Sheets.Examples.Tutorial;

public sealed class RenderExample : IExample
{
    public string Name => "tutorial/render";

    public string Description => "Render a workbook to PDF, PNG, SVG, and HTML";

    public void Run(string outputDirectory)
    {
        string source = ExampleSupport.CreateSalesWorkbook(Path.Combine(outputDirectory, "render-source.xlsx"));
        using Workbook workbook = Workbook.Open(source);

        var pdf = new PdfOptions
        {
            Title = "Q3 Report",
            Author = "Finance",
            FitTo = PdfFit.Width,
            Bookmarks = true,
        };
        workbook.SavePdf(Path.Combine(outputDirectory, "report.pdf"), pdf);

        var image = new ImageOptions { Scale = 2, Gridlines = true };
        Sheet sales = workbook.Sheets["Q3 Sales"];   // look the sheet up by name
        File.WriteAllBytes(Path.Combine(outputDirectory, "report.png"), sales.RenderImage(ImageFormat.Png, image));
        File.WriteAllBytes(Path.Combine(outputDirectory, "report.svg"), sales.RenderImage(ImageFormat.Svg, image));
        workbook.SaveHtml(Path.Combine(outputDirectory, "report.html"), new HtmlOptions { SingleFile = true });
        Console.WriteLine("Q3 Sales is " + sales.GetPageCount() + " page(s)");
    }
}
