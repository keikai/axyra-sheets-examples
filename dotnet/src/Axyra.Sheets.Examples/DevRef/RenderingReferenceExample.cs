using Axyra.Sheets.IO;

namespace Axyra.Sheets.Examples.DevRef;

public sealed class RenderingReferenceExample : IExample
{
    public string Name => "dev-ref/rendering";

    public string Description => "Exercise PDF options and all documented image/HTML render paths";

    public void Run(string outputDirectory)
    {
        string source = ExampleSupport.CreateSalesWorkbook(Path.Combine(outputDirectory, "rendering-source.xlsx"));
        using Workbook workbook = Workbook.Open(source);
        var pdf = new PdfOptions
        {
            Title = "Axyra rendering example",
            Watermark = "DRAFT",
            FitTo = PdfFit.Width,
        };
        workbook.SavePdf(Path.Combine(outputDirectory, "rendering.pdf"), pdf);

        var image = new ImageOptions { WidthInPixels = 1200, Gridlines = true };
        // By name: an evaluation-mode file has an extra first sheet.
        Sheet sales = workbook.Sheets["Q3 Sales"];
        File.WriteAllBytes(Path.Combine(outputDirectory, "rendering.png"), sales.RenderImage(ImageFormat.Png, image));
        File.WriteAllBytes(Path.Combine(outputDirectory, "rendering.svg"), sales.RenderImage(ImageFormat.Svg, image));
        File.WriteAllBytes(Path.Combine(outputDirectory, "rendering.jpg"), sales.RenderImage(ImageFormat.Jpeg, image));
        File.WriteAllBytes(Path.Combine(outputDirectory, "rendering.tiff"), sales.RenderImage(ImageFormat.Tiff, image));
        workbook.SaveHtml(Path.Combine(outputDirectory, "rendering.html"), new HtmlOptions { SingleFile = true });
    }
}
