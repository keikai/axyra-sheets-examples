using Axyra.Sheets.Style;

namespace Axyra.Sheets.Examples.FeatureDemo;

/// <summary>Rows and columns — sizing, hiding, autofit, insert and delete.</summary>
/// <remarks>
/// Row heights are in <b>points</b>, column widths in <b>character widths</b>, which is Excel's
/// own pair of units. Inserting or deleting shifts the cells below or to the right and rewrites
/// the formulas that referenced them.
/// </remarks>
public sealed class RowsColumnsExample : IExample
{
    private const string Content = "rowsColumns-data-content.xlsx";

    public string Name => "feature-demo/rows-columns";

    public string Description => "Size, autofit, hide, insert and delete rows and columns";

    public void Run(string outputDirectory)
    {
        using (Workbook wb = Sizes())
        {
            wb.Save(Path.Combine(outputDirectory, "rows-columns-sizes.xlsx"));
        }
        using (Workbook wb = Autofit())
        {
            wb.Save(Path.Combine(outputDirectory, "rows-columns-autofit.xlsx"));
        }
        using (Workbook wb = Hide())
        {
            wb.Save(Path.Combine(outputDirectory, "rows-columns-hidden.xlsx"));
        }
        using (Workbook wb = InsertDelete())
        {
            wb.Save(Path.Combine(outputDirectory, "rows-columns-insert.xlsx"));
        }
        Console.WriteLine("Wrote four row and column workbooks to " + outputDirectory);
    }

    /// <summary>Explicit sizes, and the sheet-wide defaults everything else inherits.</summary>
    private static Workbook Sizes()
    {
        Workbook wb = Data("Sizes");
        Sheet sheet = wb.Sheets[0];

        sheet.Rows[0].Height = 34;
        sheet.Columns[0].Width = 26;
        sheet.Columns[1].Width = 6;
        sheet.Rows.DefaultHeight = 18;

        int row = 8;
        sheet[row, 0].Value = "Read back";
        sheet.GetRange(row, 0, row, 0).Style = new CellStyle { Font = new Font { Bold = true } };
        Report(sheet, ++row, "Rows[0].Height", sheet.Rows[0].Height + " pt");
        Report(sheet, ++row, "Rows[0].IsHeightCustom", sheet.Rows[0].IsHeightCustom.ToString());
        Report(sheet, ++row, "Columns[0].Width", sheet.Columns[0].Width + " chars");
        Report(sheet, ++row, "Columns[2].IsWidthCustom", sheet.Columns[2].IsWidthCustom.ToString());
        Report(sheet, ++row, "Rows.DefaultHeight", sheet.Rows.DefaultHeight.ToString());
        return wb;
    }

    /// <summary>Autofit measures the rendered text, so it needs the content first.</summary>
    private static Workbook Autofit()
    {
        Workbook wb = new Workbook();
        Sheet sheet = wb.Sheets.Add("Autofit");

        sheet[0, 0].Value = "Column A was autofitted";
        sheet[1, 0].Value = "A considerably longer line of text";
        sheet[2, 0].Value = "short";
        sheet[0, 2].Value = "Column C was left at the default";
        sheet[1, 2].Value = "A considerably longer line of text";

        sheet.Columns[0].AutoFit();
        sheet.Rows[1].AutoFit();

        // Reported in G1, clear of the columns being measured: in column A it would
        // itself be the longest string there.
        sheet[0, 6].Value = "Columns[0].AutoFit() -> " + sheet.Columns[0].Width + " character widths";
        return wb;
    }

    /// <summary>Hiding keeps the data; it only sets the row or column flag.</summary>
    private static Workbook Hide()
    {
        Workbook wb = Data("Hidden");
        Sheet sheet = wb.Sheets[0];

        sheet.Rows[3].Hidden = true;
        sheet.Columns[2].Hidden = true;

        sheet[8, 0].Value = "Row 4 and column C are hidden.";
        sheet[9, 0].Value = "Rows[3].Hidden = " + sheet.Rows[3].Hidden
            + ", Columns[2].Hidden = " + sheet.Columns[2].Hidden;
        sheet[10, 0].Value = "The SUM below still counts the hidden row — hiding is presentation only.";
        sheet[11, 0].Value = "SUM(B2:B6)";
        sheet[11, 1].Formula = "SUM(B2:B6)";
        sheet.Columns[0].Width = 34;
        wb.Recalculate();
        return wb;
    }

    /// <summary>Insert and delete shift the neighbours and repair the formulas.</summary>
    private static Workbook InsertDelete()
    {
        Workbook wb = Data("Shifted");
        Sheet sheet = wb.Sheets[0];
        sheet[7, 0].Value = "Total";
        sheet[7, 1].Formula = "SUM(B2:B6)";
        sheet["A8:B8"].Style = new CellStyle { Font = new Font { Bold = true } };

        sheet.Rows.Insert(2, 2);
        sheet[2, 0].Value = "<- two rows inserted here";
        sheet["A3"].Style = new CellStyle
        {
            Font = new Font { Italic = true, Color = Color.FromRgb(0x1F, 0x4E, 0x79) },
        };

        sheet.Columns.Insert(1, 1);
        sheet[0, 1].Value = "Inserted";

        sheet[12, 0].Value = "The total's formula is now";
        sheet[13, 0].Value = sheet[9, 2].Formula;
        sheet.Columns[0].Width = 30;
        wb.Recalculate();
        return wb;
    }

    private static void Report(Sheet sheet, int row, string call, string result)
    {
        sheet[row, 0].Value = call;
        sheet[row, 1].Value = result;
    }

    /// <summary>The starting sheet: a small data table read from a packaged workbook, renamed for the sample.</summary>
    private static Workbook Data(string name)
    {
        Workbook wb = Workbook.Open(DemoBooks.Read(Content), WorkbookFormat.Xlsx);
        wb.Sheets[0].Name = name;
        return wb;
    }
}
