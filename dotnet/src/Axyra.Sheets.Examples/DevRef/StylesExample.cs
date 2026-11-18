using Axyra.Sheets.Format;
using Axyra.Sheets.Style;

namespace Axyra.Sheets.Examples.DevRef;

public sealed class StylesExample : IExample
{
    public string Name => "dev-ref/styles";

    public string Description => "Build, copy, register, and apply styles and number formats";

    public void Run(string outputDirectory)
    {
        var header = new CellStyle
        {
            Font = new Font
            {
                Name = "Calibri", Size = 11, Bold = true,
                Color = Color.FromRgb(255, 255, 255),
            },
            Fill = Color.Parse("#2C3E50")!,   // Parse returns null for text it cannot read
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center,
            WrapText = true,
            Borders = new Borders { Bottom = new BorderEdge(BorderStyle.Double, Color.Parse("#333333")) },
        };
        // Java serializes the style to JSON and back (CellStyle.toJson / fromJson). The .NET
        // CellStyle is a value record with no public JSON form: it is stored and copied as is.
        CellStyle restored = header with { };
        var money = new CellStyle { NumberFormat = "#,##0.00;[Red](#,##0.00)" };

        using Workbook workbook = new Workbook();
        Sheet sheet = workbook.Sheets.Add("Styles");
        sheet[0, 0].Value = "Amount";
        sheet[1, 0].Value = 1234.5;
        sheet[0, 0].Style = restored;
        workbook.RegisterNamedStyle("Currency", money);
        sheet["A2:A2"].ApplyNamedStyle("Currency");

        Theme theme = new Theme { MajorFont = "Georgia", MinorFont = "Calibri" }
            .WithColor(Theme.Slot.Accent1, Color.Parse("#2C3E50"));
        workbook.Theme = theme;
        ParsedValue parsed = Formats.Parse("31/12/2026");
        Console.WriteLine(Formats.Format(1234.5, "#,##0.00") + "; parsed=" + parsed);
        workbook.Save(Path.Combine(outputDirectory, "styles.xlsx"));
    }
}
