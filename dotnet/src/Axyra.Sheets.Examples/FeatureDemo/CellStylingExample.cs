using Axyra.Sheets.Style;

namespace Axyra.Sheets.Examples.FeatureDemo;

/// <summary>Cell styling — fonts, fills, borders and alignment.</summary>
/// <remarks>
/// <see cref="CellStyle"/> is an immutable value built once and applied to a whole range in one
/// call. Identical styles are interned by the engine, so styling a thousand rows with the same
/// record costs one style entry in the file, not a thousand.
/// </remarks>
public sealed class CellStylingExample : IExample
{
    private const string Content = "cellStyling-grid-content.xlsx";

    public string Name => "feature-demo/cell-styling";

    public string Description => "Style cells with fonts, fills, borders and alignment";

    public void Run(string outputDirectory)
    {
        using (Workbook wb = Fonts())
        {
            wb.Save(Path.Combine(outputDirectory, "cell-styling-fonts.xlsx"));
        }
        using (Workbook wb = Fills())
        {
            wb.Save(Path.Combine(outputDirectory, "cell-styling-fills.xlsx"));
        }
        using (Workbook wb = BordersDemo())
        {
            wb.Save(Path.Combine(outputDirectory, "cell-styling-borders.xlsx"));
        }
        using (Workbook wb = Alignment())
        {
            wb.Save(Path.Combine(outputDirectory, "cell-styling-alignment.xlsx"));
        }
        using (Workbook wb = StyledReport())
        {
            wb.Save(Path.Combine(outputDirectory, "cell-styling-report.xlsx"));
        }
        Console.WriteLine("Wrote five styling workbooks to " + outputDirectory);
    }

    private static Workbook Fonts()
    {
        Workbook wb = Grid("Fonts");
        Sheet sheet = wb.Sheets[0];

        Label(sheet, 1, "Font.Name = \"Georgia\"");
        sheet["B2"].Style = new CellStyle { Font = new Font { Name = "Georgia", Size = 13 } };

        Label(sheet, 2, "Font.Size = 18");
        sheet["B3"].Style = new CellStyle { Font = new Font { Size = 18 } };

        Label(sheet, 3, "Bold + Italic");
        sheet["B4"].Style = new CellStyle { Font = new Font { Bold = true, Italic = true } };

        Label(sheet, 4, "Underline.Double");
        sheet["B5"].Style = new CellStyle { Font = new Font { Underline = Underline.Double } };

        Label(sheet, 5, "Font.Color (dark red)");
        sheet["B6"].Style = new CellStyle { Font = new Font { Color = Color.FromRgb(0xC0, 0x00, 0x00) } };

        Label(sheet, 6, "Strikethrough = true");
        sheet["B7"].Style = new CellStyle { Font = new Font { Strikethrough = true } };
        return wb;
    }

    private static Workbook Fills()
    {
        Workbook wb = Grid("Fills");
        Sheet sheet = wb.Sheets[0];

        Label(sheet, 1, "Fill.Solid");
        sheet["B2"].Style = new CellStyle { Fill = Fill.Solid(Color.FromRgb(0xDD, 0xEB, 0xF7)) };

        Label(sheet, 2, "Fill.Solid, any colour");
        sheet["B3"].Style = new CellStyle { Fill = Fill.Solid(Color.FromRgb(0xFF, 0xF2, 0xCC)) };

        Label(sheet, 3, "Fill.Solid + Font.Color");
        sheet["B4"].Style = new CellStyle
        {
            Fill = Fill.Solid(Color.FromRgb(0x1F, 0x4E, 0x79)),
            Font = new Font { Color = Color.White },
        };

        sheet.Columns[0].Width = 34;
        return wb;
    }

    private static Workbook BordersDemo()
    {
        Workbook wb = Grid("Borders");
        Sheet sheet = wb.Sheets[0];

        Label(sheet, 1, "per edge, per style");
        sheet["B2"].Style = new CellStyle
        {
            Borders = new Borders
            {
                Top = new BorderEdge(BorderStyle.Medium, Color.Black),
                Bottom = new BorderEdge(BorderStyle.Double, Color.FromRgb(0xC0, 0x00, 0x00)),
                Left = new BorderEdge(BorderStyle.Thin, Color.Gray),
                Right = new BorderEdge(BorderStyle.Dashed, Color.Gray),
            },
        };

        Label(sheet, 3, "Borders.Around over a block ->");
        sheet["D4:G9"].Style = new CellStyle
        {
            Borders = Borders.Around(BorderStyle.Thin, Color.FromRgb(0x1F, 0x4E, 0x79)),
        };

        sheet.Columns[0].Width = 34;
        return wb;
    }

    private static Workbook Alignment()
    {
        Workbook wb = Grid("Alignment");
        Sheet sheet = wb.Sheets[0];
        sheet.Rows[1].Height = 36;
        sheet.Rows[2].Height = 36;
        sheet.Rows[3].Height = 46;

        Label(sheet, 1, "Center / Center");
        sheet["B2"].Style = new CellStyle
        {
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center,
        };

        Label(sheet, 2, "Right / Bottom");
        sheet["B3"].Style = new CellStyle
        {
            HorizontalAlignment = HorizontalAlignment.Right,
            VerticalAlignment = VerticalAlignment.Bottom,
        };

        sheet[3, 0].Value = "WrapText + Indent";
        sheet[3, 1].Value = "A longer sentence that has to wrap inside its own cell.";
        sheet["B4"].Style = new CellStyle
        {
            WrapText = true,
            VerticalAlignment = VerticalAlignment.Top,
            Indent = 1,
        };
        return wb;
    }

    /// <summary>The everyday case: one header style, one body style, applied by range.</summary>
    private static Workbook StyledReport()
    {
        Workbook wb = new Workbook();
        Sheet sheet = wb.Sheets.Add("Report");

        CellStyle header = new CellStyle
        {
            Font = new Font { Bold = true, Color = Color.White },
            Fill = Fill.Solid(Color.FromRgb(0x1F, 0x4E, 0x79)),
            HorizontalAlignment = HorizontalAlignment.Center,
            Borders = Borders.Around(BorderStyle.Thin, Color.FromRgb(0x1F, 0x4E, 0x79)),
        };
        CellStyle money = new CellStyle
        {
            NumberFormat = "#,##0.00",
            Borders = Borders.Around(BorderStyle.Hair, Color.FromRgb(0xBF, 0xBF, 0xBF)),
        };

        sheet["A1:D1"].SetValues(new CellValue[,] { { "Client", "Hours", "Rate", "Fee" } });
        string[] clients = ["Adler", "Brandt", "Cortez", "Duval", "Eriksen"];
        double[] hours = [42, 17.5, 63, 8, 29.25];
        double[] rate = [145, 190, 120, 210, 160];
        for (int i = 0; i < clients.Length; i++)
        {
            sheet[i + 1, 0].Value = clients[i];
            sheet[i + 1, 1].Value = hours[i];
            sheet[i + 1, 2].Value = rate[i];
            sheet[i + 1, 3].Formula = "B" + (i + 2) + "*C" + (i + 2);
        }

        sheet["A1:D1"].Style = header;
        sheet.GetRange(1, 1, clients.Length, 3).Style = money;
        sheet.Columns[0].Width = 16;
        sheet.Columns[3].Width = 14;
        wb.Recalculate();
        return wb;
    }

    /// <summary>The starting sheet: a labelled grid read from a packaged workbook, renamed for the sample.</summary>
    private static Workbook Grid(string name)
    {
        Workbook wb = Workbook.Open(DemoBooks.Read(Content), WorkbookFormat.Xlsx);
        wb.Sheets[0].Name = name;
        return wb;
    }

    private static void Label(Sheet sheet, int row, string call)
    {
        sheet[row, 0].Value = call;
        sheet[row, 1].Value = "Sample text 1234";
    }
}
