using Axyra.Sheets.Charts;
using Axyra.Sheets.Drawing;
using Axyra.Sheets.IO;
using Axyra.Sheets.Style;

namespace Axyra.Sheets.Examples.FeatureDemo;

/// <summary>
/// Export to PDF.
/// </summary>
/// <remarks>
/// <see cref="Workbook.ToPdf"/> returns the file as a <c>byte[]</c>, which is what a
/// server-side report actually needs — straight into a download, an email attachment or a
/// blob column, with no temporary file. <see cref="PdfOptions"/> selects sheets, a
/// printed-page range, and how the content is fitted; <see cref="Sheet.GetPageCount"/>
/// says in advance how many pages it will be.
/// </remarks>
public sealed class ExportPdfExample : IExample
{
    public string Name => "feature-demo/export-pdf";

    public string Description => "Render a workbook to PDF bytes: whole book, one sheet, fitted, and page counts";

    public void Run(string outputDirectory)
    {
        using (Workbook wb = Report())
        {
            wb.Save(Path.Combine(outputDirectory, "export-pdf-report.xlsx"));
        }
        Write(Path.Combine(outputDirectory, "report.pdf"), Whole());
        Write(Path.Combine(outputDirectory, "report-sheet1.pdf"), FirstSheet());
        Write(Path.Combine(outputDirectory, "report-fit-width.pdf"), Fitted());
        using (Workbook wb = PageCount())
        {
            wb.Save(Path.Combine(outputDirectory, "export-pdf-page-count.xlsx"));
        }
    }

    private static void Write(string file, byte[] pdf)
    {
        File.WriteAllBytes(file, pdf);
        Console.WriteLine("Wrote " + Path.GetFileName(file) + " (" + pdf.Length + " bytes)");
    }

    /// <summary>Every sheet, with the default options.</summary>
    private static byte[] Whole()
    {
        using Workbook wb = Report();
        return wb.ToPdf(new PdfOptions());
    }

    /// <summary><c>Sheets</c> takes indexes; <c>SheetNames</c> takes names.</summary>
    private static byte[] FirstSheet()
    {
        using Workbook wb = Report();
        return wb.ToPdf(new PdfOptions { Sheets = [0] });
    }

    /// <summary>Squeeze a wide sheet onto one page across.</summary>
    private static byte[] Fitted()
    {
        using Workbook wb = Report();
        return wb.ToPdf(new PdfOptions { FitTo = PdfFit.Width });
    }

    /// <summary>How long the PDF will be, before rendering it.</summary>
    private static Workbook PageCount()
    {
        Workbook wb = Report();
        Sheet sheet = wb.Sheets[0];
        int row = 30;   // clear of the chart, which occupies rows 12-29

        // Java has wb.sheetPageCount(i) and Renderer.pageCount(wb, i); .NET has the one
        // member, Sheet.GetPageCount(), so the third line asks it again.
        sheet[row, 0].Value = "Pages — Summary";
        sheet[row, 3].Value = wb.Sheets[0].GetPageCount();
        sheet[row + 1, 0].Value = "Pages — By rep";
        sheet[row + 1, 3].Value = wb.Sheets[1].GetPageCount();
        sheet[row + 2, 0].Value = "Renderer.pageCount(wb, 0)";
        sheet[row + 2, 3].Value = wb.Sheets[0].GetPageCount();
        sheet[row + 3, 0].Value = "Whole workbook prints both; sheets(0) prints the first alone.";
        sheet.GetRange(row, 0, row + 2, 0).Style = new CellStyle { Font = new() { Bold = true } };
        sheet.Columns[0].Width = 30;
        return wb;
    }

    /// <summary>A report worth printing: a title band, figures, a total and a chart.</summary>
    private static Workbook Report()
    {
        Workbook wb = new Workbook();
        Sheet sheet = wb.Sheets.Add("Summary");

        sheet[0, 0].Value = "Regional performance — Q3 2026";
        sheet["A1:D1"].Merge();
        sheet["A1:D1"].Style = TitleBand();
        sheet.Rows[0].Height = 30;

        sheet["A3:D3"].SetValues(new CellValue[,] { { "Region", "Units", "Revenue", "Margin" } });
        sheet["A3:D3"].Style = new CellStyle { Font = new() { Bold = true } };

        string[] regions = ["North", "South", "East", "West", "Central"];
        double[] units = [1240, 980, 1610, 1130, 870];
        double[] revenue = [18600, 14700, 24150, 16950, 13050];
        for (int i = 0; i < regions.Length; i++)
        {
            sheet[i + 3, 0].Value = regions[i];
            sheet[i + 3, 1].Value = units[i];
            sheet[i + 3, 2].Value = revenue[i];
            sheet[i + 3, 3].Formula = "C" + (i + 4) + "/B" + (i + 4);
        }

        int totalRow = regions.Length + 3;
        sheet[totalRow, 0].Value = "Total";
        sheet[totalRow, 1].Formula = "SUM(B4:B8)";
        sheet[totalRow, 2].Formula = "SUM(C4:C8)";
        sheet.GetRange(totalRow, 0, totalRow, 3).Style = new CellStyle { Font = new() { Bold = true } };
        sheet.GetRange(3, 2, totalRow, 2).Style = new CellStyle { NumberFormat = "#,##0" };
        sheet.GetRange(3, 3, totalRow - 1, 3).Style = new CellStyle { NumberFormat = "#,##0.00" };
        sheet.Columns[0].Width = 16;
        sheet.Columns[2].Width = 14;

        // Corner cells, as Java's anchor(11, 0, 28, 6): the chart ends at the top-left of G29.
        Chart chart = sheet.Charts.Add(ChartType.Column, new DrawingAnchor(11, 0, 28, 6));
        chart.Title = "Revenue by region";
        chart.Series.Add("Revenue", "Summary!C4:C8", "Summary!A4:A8");
        chart.Legend.Visible = false;

        // The summary prints the table and the chart below it, which ends in row 29.
        sheet.PageSetup.PrintArea = "A1:G29";

        Detail(wb);
        wb.Recalculate();
        return wb;
    }

    /// <summary>
    /// A second sheet, long enough to be a page of its own, so that
    /// <c>PdfOptions.Sheets</c> has something to select.
    /// </summary>
    private static void Detail(Workbook wb)
    {
        Sheet sheet = wb.Sheets.Add("By rep");

        sheet[0, 0].Value = "Revenue by representative — Q3 2026";
        sheet["A1:C1"].Merge();
        sheet["A1:C1"].Style = TitleBand();
        sheet.Rows[0].Height = 30;

        sheet["A3:C3"].SetValues(new CellValue[,] { { "Rep", "Region", "Revenue" } });
        sheet["A3:C3"].Style = new CellStyle { Font = new() { Bold = true } };

        string[] reps = ["Adler", "Brandt", "Cortez", "Duval", "Eriksen",
                "Fischer", "Goddard", "Halvorsen", "Ibarra", "Jansen"];
        string[] regions = ["North", "South", "East", "West", "Central"];
        int firstData = 3;
        for (int i = 0; i < reps.Length; i++)
        {
            sheet[firstData + i, 0].Value = reps[i];
            sheet[firstData + i, 1].Value = regions[i % regions.Length];
            sheet[firstData + i, 2].Value = 2600 + (i * 737) % 4200;
        }

        int lastData = firstData + reps.Length - 1;
        int totalRow = lastData + 2;
        sheet[totalRow, 0].Value = "Total";
        sheet[totalRow, 2].Formula = "SUM(C" + (firstData + 1) + ":C" + (lastData + 1) + ")";
        sheet.GetRange(firstData, 2, totalRow, 2).Style = new CellStyle { NumberFormat = "#,##0" };
        sheet.GetRange(totalRow, 0, totalRow, 2).Style = new CellStyle
        {
            Font = new() { Bold = true },
            NumberFormat = "#,##0",
        };
        sheet.Columns[0].Width = 18;
        sheet.Columns[1].Width = 14;
        sheet.Columns[2].Width = 14;
        sheet.PageSetup.PrintArea = "A1:C" + (totalRow + 1);
    }

    private static CellStyle TitleBand() => new()
    {
        Font = new() { Bold = true, Size = 15, Color = Color.White },
        Fill = Color.FromRgb(0x1F, 0x4E, 0x79),
        HorizontalAlignment = HorizontalAlignment.Center,
    };
}
