using Axyra.Sheets.Style;

namespace Axyra.Sheets.Examples.FeatureDemo;

/// <summary>Create a new workbook.</summary>
/// <remarks>
/// <c>new Workbook()</c> builds an empty book on the sparse backend, the general-purpose
/// default. <c>new Workbook(WorkbookStorage.Columnar)</c> builds a behaviourally identical one
/// on columnar storage, which is far leaner in memory for large, dense, numeric sheets. A book
/// read from a file always uses the reader's representation, so the choice only applies to
/// books you generate.
/// </remarks>
public sealed class CreateWorkbookExample : IExample
{
    public string Name => "feature-demo/create-workbook";

    public string Description => "Create workbooks on the sparse and columnar backends, and set workbook defaults";

    public void Run(string outputDirectory)
    {
        using (Workbook wb = Sparse())
        {
            wb.Save(Path.Combine(outputDirectory, "create-workbook-sparse.xlsx"));
        }
        using (Workbook wb = Columnar())
        {
            wb.Save(Path.Combine(outputDirectory, "create-workbook-columnar.xlsx"));
        }
        using (Workbook wb = Defaults())
        {
            wb.Save(Path.Combine(outputDirectory, "create-workbook-defaults.xlsx"));
        }
        Console.WriteLine("Created three workbooks in " + outputDirectory);
    }

    /// <summary>The general-purpose backend: a book from nothing, with one cell in it.</summary>
    private static Workbook Sparse()
    {
        Workbook wb = new Workbook();
        Sheet sheet = wb.Sheets.Add("Sheet1");
        sheet[0, 0].Value = "Hello";

        sheet[2, 0].Value = "new Workbook() — sparse storage, the general-purpose default.";
        sheet.Columns[0].Width = 60;
        return wb;
    }

    /// <summary>
    /// The columnar backend trades random-access cost for a much smaller footprint on wide,
    /// dense, numeric sheets.
    /// </summary>
    private static Workbook Columnar()
    {
        Workbook wb = new Workbook(WorkbookStorage.Columnar);
        Sheet sheet = wb.Sheets.Add("Wide");
        sheet["A1:Z1"].SetNumbers(new double[26]);

        sheet[2, 0].Value = "new Workbook(WorkbookStorage.Columnar) — 26 columns written in one call.";
        sheet[3, 0].Value = "Same API and same behaviour as new Workbook(); leaner at rest on large, "
            + "dense, numeric sheets.";
        sheet.Columns[0].Width = 60;
        return wb;
    }

    /// <summary>Workbook-wide defaults, written onto the sheet rather than to the console.</summary>
    private static Workbook Defaults()
    {
        Workbook wb = new Workbook();
        Sheet sheet = wb.Sheets.Add("Defaults");

        wb.DefaultFont = new Font { Name = "Calibri", Size = 11 };

        sheet["A1:B1"].SetValues(new CellValue[,] { { "Property", "Value" } });
        sheet["A1:B1"].Style = new CellStyle { Font = new Font { Bold = true } };

        sheet[1, 0].Value = "default font";
        sheet[1, 1].Value = wb.DefaultStyle.Font.Name + " " + wb.DefaultStyle.Font.Size;
        sheet[2, 0].Value = "locale";
        sheet[2, 1].Value = wb.Locale;
        sheet[3, 0].Value = "1904 date system";
        sheet[3, 1].Value = wb.Date1904;
        sheet[4, 0].Value = "sheets";
        sheet[4, 1].Value = wb.Sheets.Count;

        sheet.Columns[0].Width = 24;
        sheet.Columns[1].Width = 24;
        return wb;
    }
}
