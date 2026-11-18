using Axyra.Sheets.IO;

namespace Axyra.Sheets.Examples.Guides;

/// <summary>The complete program from "How to Convert Excel to PDF in .NET".</summary>
public sealed class ExcelToPdfExample : IExample
{
    public string Name => "guides/convert-excel-to-pdf";

    public string Description => "Convert every visible worksheet to a fitted, bookmarked PDF";

    public void Run(string outputDirectory)
    {
        string source = ExampleSupport.CreateSalesWorkbook(Path.Combine(outputDirectory, "pdf-source.xlsx"));
        string target = Path.Combine(outputDirectory, "report.pdf");

        using Workbook wb = Workbook.Open(source);
        var worksheets = new List<string>();

        foreach (Sheet sheet in wb.Sheets)
        {
            // Naming a hidden sheet would render it, so skip those too.
            if (sheet.SheetType != SheetType.Worksheet || sheet.Visibility != Visibility.Visible)
            {
                continue;
            }
            worksheets.Add(sheet.Name);

            PageSetup page = sheet.PageSetup;
            page.Orientation = PageOrientation.Landscape;
            page.PaperSize = 9;                                   // A4
            page.Margins = new PageMargins(0.5, 0.5, 0.75, 0.75, 0.3, 0.3);
            page.PrintTitleRows = "$1:$1";                        // header on every page
        }

        var options = new PdfOptions
        {
            Title = "Quarterly Report",
            Author = "Finance",
            SheetNames = worksheets,
            FitTo = PdfFit.Width,
            Bookmarks = true,
        };

        wb.SavePdf(target, options);

        string first = worksheets[0];
        int pages = wb.Sheets[first].GetPageCount();
        Console.WriteLine("wrote " + target + "; " + first + " is " + pages + " page(s)");
    }
}
