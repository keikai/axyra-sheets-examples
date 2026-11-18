using Axyra.Sheets.Style;

namespace Axyra.Sheets.Examples.FeatureDemo;

/// <summary>Merge cells and freeze panes — the two things that make a long sheet readable.</summary>
/// <remarks>
/// <c>Merge()</c> joins a rectangle into one cell; <c>AutoMerge()</c> does it for you wherever
/// consecutive cells repeat a value, which is what a grouped report usually wants.
/// <c>FreezePanes(rows, columns)</c> pins the headings so they stay put while the body scrolls.
/// </remarks>
public sealed class MergeFreezeExample : IExample
{
    public string Name => "feature-demo/merge-freeze";

    public string Description => "Merge a title band, auto-merge repeated values and freeze panes";

    public void Run(string outputDirectory)
    {
        using (Workbook wb = TitleBand())
        {
            wb.Save(Path.Combine(outputDirectory, "merge-freeze-title-band.xlsx"));
        }
        using (Workbook wb = AutoMerge())
        {
            wb.Save(Path.Combine(outputDirectory, "merge-freeze-auto-merge.xlsx"));
        }
        using (Workbook wb = ListMerges())
        {
            wb.Save(Path.Combine(outputDirectory, "merge-freeze-list-merges.xlsx"));
        }
        using (Workbook wb = Freeze())
        {
            wb.Save(Path.Combine(outputDirectory, "merge-freeze-frozen.xlsx"));
        }
        Console.WriteLine("Wrote four merge and freeze workbooks to " + outputDirectory);
    }

    /// <summary>The classic report header: one merged band across the columns.</summary>
    private static Workbook TitleBand()
    {
        Workbook wb = new Workbook();
        Sheet sheet = Ledger(wb, "Titled", 12);

        sheet.Rows.Insert(0, 2);
        sheet[0, 0].Value = "Quarterly ledger — Q3 2026";
        sheet["A1:D1"].Merge();
        sheet["A1:D1"].Style = new CellStyle
        {
            Font = new Font { Bold = true, Size = 15, Color = Color.White },
            Fill = Fill.Solid(Color.FromRgb(0x1F, 0x4E, 0x79)),
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center,
        };
        sheet.Rows[0].Height = 30;
        return wb;
    }

    /// <summary>Runs of equal values collapse into one merged cell each.</summary>
    private static Workbook AutoMerge()
    {
        Workbook wb = new Workbook();
        Sheet sheet = Ledger(wb, "Grouped", 12);

        int merged = sheet.GetRange(1, 0, 12, 0).AutoMerge(false);
        sheet.GetRange(1, 0, 12, 0).Style = new CellStyle
        {
            VerticalAlignment = VerticalAlignment.Center,
            HorizontalAlignment = HorizontalAlignment.Center,
            Borders = Borders.Around(BorderStyle.Thin, Color.FromRgb(0xBF, 0xBF, 0xBF)),
        };

        sheet[14, 0].Value = "AutoMerge(false) created " + merged + " vertical regions in column A.";
        sheet.Columns[0].Width = 16;
        return wb;
    }

    /// <summary>Every merged rectangle is readable back as a <see cref="CellRange"/>.</summary>
    private static Workbook ListMerges()
    {
        Workbook wb = AutoMerge();
        Sheet sheet = wb.Sheets[0];

        int row = 16;
        sheet[row, 0].Value = "Merged regions";
        sheet.GetRange(row, 0, row, 0).Style = new CellStyle { Font = new Font { Bold = true } };
        foreach (CellRange region in sheet.MergedRegions)
        {
            row++;
            sheet[row, 0].Value = "r" + (region.FirstRow + 1) + "c" + (region.FirstColumn + 1)
                + " to r" + (region.LastRow + 1) + "c" + (region.LastColumn + 1);
        }
        return wb;
    }

    /// <summary>The heading row and the label column stay put while the rest scrolls.</summary>
    private static Workbook Freeze()
    {
        Workbook wb = new Workbook();
        Sheet sheet = Ledger(wb, "Frozen", 60);

        sheet.FreezePanes(1, 1);
        sheet[0, 4].Value = "FreezePanes(1, 1) — scroll down, the header row stays.";
        sheet.Columns[4].Width = 46;
        return wb;
    }

    /// <summary>A ledger whose first column repeats, so grouping has something to do.</summary>
    private static Sheet Ledger(Workbook wb, string name, int rows)
    {
        Sheet sheet = wb.Sheets.Add(name);
        sheet["A1:D1"].SetValues(new CellValue[,] { { "Region", "Month", "Units", "Revenue" } });
        sheet["A1:D1"].Style = new CellStyle { Font = new Font { Bold = true } };

        string[] regions = ["North", "South", "East", "West"];
        string[] months = ["Jul", "Aug", "Sep"];
        for (int i = 0; i < rows; i++)
        {
            sheet[i + 1, 0].Value = regions[(i / 3) % regions.Length];
            sheet[i + 1, 1].Value = months[i % months.Length];
            sheet[i + 1, 2].Value = 80 + (i * 37) % 220;
            sheet[i + 1, 3].Value = 1200 + (i * 613) % 4300;
        }
        sheet.GetRange(1, 3, rows, 3).Style = new CellStyle { NumberFormat = "#,##0" };
        sheet.Columns[0].Width = 14;
        sheet.Columns[3].Width = 14;
        return sheet;
    }
}
