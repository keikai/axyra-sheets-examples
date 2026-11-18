using Axyra.Sheets.Style;

namespace Axyra.Sheets.Examples.FeatureDemo;

/// <summary>Manage worksheets — create, insert, rename, reorder, copy, delete.</summary>
/// <remarks>
/// Renaming rewrites every formula that referenced the old name, and deleting follows Excel's
/// own activation rules.
/// </remarks>
public sealed class ManageWorksheetsExample : IExample
{
    public string Name => "feature-demo/manage-worksheets";

    public string Description => "Create, insert, rename, reorder, copy, delete and colour worksheets";

    public void Run(string outputDirectory)
    {
        using (Workbook wb = CreateAndInsert())
        {
            wb.Save(Path.Combine(outputDirectory, "manage-worksheets-create-insert.xlsx"));
        }
        using (Workbook wb = RenameAndReorder())
        {
            wb.Save(Path.Combine(outputDirectory, "manage-worksheets-rename-reorder.xlsx"));
        }
        using (Workbook wb = CopySheet())
        {
            wb.Save(Path.Combine(outputDirectory, "manage-worksheets-copy.xlsx"));
        }
        using (Workbook wb = DeleteSheet())
        {
            wb.Save(Path.Combine(outputDirectory, "manage-worksheets-delete.xlsx"));
        }
        using (Workbook wb = ColourTabs())
        {
            wb.Save(Path.Combine(outputDirectory, "manage-worksheets-tab-colours.xlsx"));
        }
        Console.WriteLine("Wrote five worksheet-management workbooks to " + outputDirectory);
    }

    /// <summary>Append with <c>Sheets.Add</c>, or place one with <c>Sheets.Insert</c>.</summary>
    private static Workbook CreateAndInsert()
    {
        Workbook wb = new Workbook();
        wb.Sheets.Add("Data");
        wb.Sheets.Add("Report");
        Sheet cover = wb.Sheets.Insert(0, "Cover");
        Label(cover, "Cover", "Sheets.Insert(0, \"Cover\") put this tab first.");
        Label(wb.Sheets["Data"], "Data", "Sheets.Add appends at the end.");
        Label(wb.Sheets["Report"], "Report", "Three sheets, in creation order.");
        wb.ActiveSheet = cover;
        return wb;
    }

    /// <summary>A rename rewrites the formulas that pointed at the old name.</summary>
    private static Workbook RenameAndReorder()
    {
        Workbook wb = new Workbook();
        Sheet data = wb.Sheets.Add("Data");
        Sheet report = wb.Sheets.Add("Report");
        data["A1:A4"].SetNumbers([10, 20, 30, 40]);

        Label(report, "Report", "This total references the other sheet.");
        report[2, 0].Value = "Total";
        report[2, 1].Formula = "SUM(Data!A1:A4)";

        // The formula above now reads SUM(SourceData!A1:A4).
        data.Name = "SourceData";
        report[4, 0].Value = "Formula after the rename:";
        report[5, 0].Value = report[2, 1].Formula;

        wb.Sheets.Move(report.Index, 0);
        wb.Recalculate();
        wb.ActiveSheet = wb.Sheets[0];
        return wb;
    }

    /// <summary>A deep copy: values, styles and formulas are re-interned into the target.</summary>
    private static Workbook CopySheet()
    {
        Workbook wb = new Workbook();
        Sheet source = wb.Sheets.Add("Template");
        Label(source, "Template", "Styled once, then copied.");
        source["A4:C4"].SetValues(new CellValue[,] { { "Item", "Qty", "Price" } });
        source["A4:C4"].Style = new CellStyle
        {
            Font = new Font { Bold = true },
            Fill = Color.FromRgb(0xDD, 0xEB, 0xF7),
        };

        wb.Sheets.AddCopy(source, "January");
        wb.Sheets.AddCopy(source, "February");
        wb.Sheets.AddCopy(source, "March");
        return wb;
    }

    /// <summary>Delete by index; the surviving tabs shift down.</summary>
    private static Workbook DeleteSheet()
    {
        Workbook wb = new Workbook();
        wb.Sheets.Add("Keep");
        Sheet scratch = wb.Sheets.Add("Scratch");
        wb.Sheets.Add("Also keep");

        Label(wb.Sheets["Keep"], "Keep", "Three sheets were created.");
        wb.Sheets.RemoveAt(scratch.Index);
        Label(wb.Sheets["Also keep"], "Also keep",
            "\"Scratch\" was deleted — " + wb.Sheets.Count + " sheets remain.");
        return wb;
    }

    /// <summary>Tab colour is a sheet property, not a style.</summary>
    private static Workbook ColourTabs()
    {
        Workbook wb = new Workbook();
        string[] names = ["Q1", "Q2", "Q3", "Q4"];
        Color[] colors =
        [
            Color.FromRgb(0x1F, 0x4E, 0x79), Color.FromRgb(0x2E, 0x75, 0xB6),
            Color.FromRgb(0x63, 0xBE, 0x7B), Color.FromRgb(0xC0, 0x00, 0x00),
        ];
        for (int i = 0; i < names.Length; i++)
        {
            Sheet sheet = wb.Sheets.Add(names[i]);
            sheet.TabColor = colors[i];
            Label(sheet, names[i], "TabColor tints the sheet tab.");
        }
        return wb;
    }

    private static void Label(Sheet sheet, string heading, string note)
    {
        sheet[0, 0].Value = heading;
        sheet["A1"].Style = new CellStyle { Font = new Font { Bold = true, Size = 14 } };
        sheet[1, 0].Value = note;
        sheet.Columns[0].Width = 22;
    }
}
