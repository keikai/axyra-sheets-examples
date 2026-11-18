using Axyra.Sheets.Style;

namespace Axyra.Sheets.Examples.FeatureDemo;

/// <summary>Save and export — one workbook, several formats, no temporary files.</summary>
/// <remarks>
/// <c>ToByteArray(format)</c> serializes straight to a <c>byte[]</c>, which is what a web
/// application actually wants: the same call feeds a download, a blob column or an HTTP
/// response. <c>Save(path)</c> picks the writer from the file extension instead.
/// </remarks>
public sealed class SaveExportExample : IExample
{
    private static readonly WorkbookFormat[] Formats =
        [WorkbookFormat.Xlsx, WorkbookFormat.Csv, WorkbookFormat.Json, WorkbookFormat.Html];

    public string Name => "feature-demo/save-export";

    public string Description => "Serialize one workbook to xlsx, csv, json and html bytes";

    public void Run(string outputDirectory)
    {
        foreach (WorkbookFormat format in Formats)
        {
            byte[] bytes = Export(format);
            string file = Path.Combine(outputDirectory, "quarterly-report." + format.Extension());
            File.WriteAllBytes(file, bytes);
            Console.WriteLine("Wrote " + Path.GetFileName(file) + " (" + bytes.Length + " bytes)");
        }
    }

    /// <summary>Serialize to <paramref name="format"/>, ready for a download, an attachment or a blob column.</summary>
    private static byte[] Export(WorkbookFormat format)
    {
        using Workbook wb = Report();
        return wb.ToByteArray(format);
    }

    /// <summary>A small report worth exporting: labels, numbers and a live total.</summary>
    private static Workbook Report()
    {
        Workbook wb = new Workbook();
        Sheet sheet = wb.Sheets.Add("Q3");

        sheet["A1:C1"].SetValues(new CellValue[,] { { "Region", "Units", "Revenue" } });
        sheet["A1:C1"].Style = new CellStyle { Font = new Font { Bold = true } };

        string[] regions = ["North", "South", "East", "West"];
        double[] units = [1240, 980, 1610, 1130];
        double[] revenue = [18600, 14700, 24150, 16950];
        for (int i = 0; i < regions.Length; i++)
        {
            sheet[i + 1, 0].Value = regions[i];
            sheet[i + 1, 1].Value = units[i];
            sheet[i + 1, 2].Value = revenue[i];
        }

        sheet[5, 0].Value = "Total";
        sheet[5, 1].Formula = "SUM(B2:B5)";
        sheet[5, 2].Formula = "SUM(C2:C5)";
        sheet["A6:C6"].Style = new CellStyle { Font = new Font { Bold = true } };
        sheet["C2:C6"].Style = new CellStyle { NumberFormat = "#,##0" };
        sheet.Columns[0].Width = 14;
        sheet.Columns[2].Width = 14;

        // CSV, JSON and HTML all read the calculated values, so evaluate first.
        wb.Recalculate();
        return wb;
    }
}
