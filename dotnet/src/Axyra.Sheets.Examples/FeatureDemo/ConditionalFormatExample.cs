using Axyra.Sheets.Content;
using Axyra.Sheets.Style;

namespace Axyra.Sheets.Examples.FeatureDemo;

/// <summary>
/// Conditional formatting — styling driven by the values themselves.
/// </summary>
/// <remarks>
/// A rule is attached to a range and re-evaluated by the spreadsheet application, so
/// editing a cell restyles it without any .NET code running. Rules stack: several on one
/// range apply in order, which is how a heat map and a threshold flag coexist.
/// </remarks>
public sealed class ConditionalFormatExample : IExample
{
    private const string Content = "conditionalFormat-figures-content.xlsx";
    private const int Rows = 18;

    public string Name => "feature-demo/conditional-format";

    public string Description => "Add threshold, data bar, colour scale, icon set and formula rules";

    public void Run(string outputDirectory)
    {
        using (Workbook wb = Threshold())
        {
            wb.Save(Path.Combine(outputDirectory, "conditional-format-threshold.xlsx"));
        }
        using (Workbook wb = DataBars())
        {
            wb.Save(Path.Combine(outputDirectory, "conditional-format-data-bars.xlsx"));
        }
        using (Workbook wb = ColorScale())
        {
            wb.Save(Path.Combine(outputDirectory, "conditional-format-color-scale.xlsx"));
        }
        using (Workbook wb = IconSet())
        {
            wb.Save(Path.Combine(outputDirectory, "conditional-format-icon-set.xlsx"));
        }
        using (Workbook wb = RowRule())
        {
            wb.Save(Path.Combine(outputDirectory, "conditional-format-row-rule.xlsx"));
        }
        Console.WriteLine("Wrote five conditional-formatting workbooks to " + outputDirectory);
    }

    /// <summary>Two stacked rules on one column: over target, and under water.</summary>
    private static Workbook Threshold()
    {
        Workbook wb = Figures();
        Sheet sheet = wb.Sheets[0];

        sheet.ConditionalFormats.Add(sheet.GetRange(1, 3, Rows, 3),
                new CellValueRule(ComparisonOperator.GreaterThan, "3000")
                {
                    Style = new ConditionalStyle
                    {
                        FillColor = Color.FromRgb(0xC6, 0xEF, 0xCE),
                        FontColor = Color.FromRgb(0x00, 0x61, 0x00),
                        Bold = true,
                    },
                });

        sheet.ConditionalFormats.Add(sheet.GetRange(1, 4, Rows, 4),
                new CellValueRule(ComparisonOperator.LessThan, "0")
                {
                    Style = new ConditionalStyle
                    {
                        FillColor = Color.FromRgb(0xFF, 0xC7, 0xCE),
                        FontColor = Color.FromRgb(0x9C, 0x00, 0x06),
                    },
                });

        Note(sheet, "Revenue over 3,000 is green; a negative variance is red. "
                + "Edit a cell and the styling follows.");
        return wb;
    }

    /// <summary>In-cell bars, proportional to the value.</summary>
    private static Workbook DataBars()
    {
        Workbook wb = Figures();
        Sheet sheet = wb.Sheets[0];

        sheet.ConditionalFormats.Add(sheet.GetRange(1, 2, Rows, 2),
                new DataBarRule(Color.FromRgb(0x63, 0x8E, 0xC6))
                {
                    Gradient = true,
                    BorderColor = Color.FromRgb(0x1F, 0x4E, 0x79),
                    ShowValue = true,
                });

        sheet.ConditionalFormats.Add(sheet.GetRange(1, 4, Rows, 4),
                new DataBarRule(Color.FromRgb(0x63, 0xBE, 0x7B))
                {
                    NegativeFillColor = Color.FromRgb(0xC0, 0x00, 0x00),
                    AxisColor = Color.Black,
                    AxisPosition = DataBarAxisPosition.Middle,
                    ShowValue = true,
                });

        Note(sheet, "Units get a plain bar; variance gets a midpoint axis so "
                + "negatives run left.");
        return wb;
    }

    /// <summary>A heat map across the revenue column.</summary>
    private static Workbook ColorScale()
    {
        Workbook wb = Figures();
        Sheet sheet = wb.Sheets[0];

        sheet.ConditionalFormats.Add(sheet.GetRange(1, 3, Rows, 3),
                ColorScaleRule.ThreeColor(Color.FromRgb(0xF8, 0x69, 0x6B),
                        Color.FromRgb(0xFF, 0xEB, 0x84),
                        Color.FromRgb(0x63, 0xBE, 0x7B)));

        Note(sheet, "A three-colour scale interpolates between the extremes of the range.");
        return wb;
    }

    /// <summary>Traffic lights, bucketed by the engine.</summary>
    private static Workbook IconSet()
    {
        Workbook wb = Figures();
        Sheet sheet = wb.Sheets[0];

        sheet.ConditionalFormats.Add(sheet.GetRange(1, 3, Rows, 3),
                new IconSetRule(IconSetType.ThreeTrafficLights1));

        Note(sheet, "iconSet(\"3TrafficLights1\") — the thresholds default to thirds.");
        return wb;
    }

    /// <summary>A formula rule styles the whole row from one column's value.</summary>
    private static Workbook RowRule()
    {
        Workbook wb = Figures();
        Sheet sheet = wb.Sheets[0];

        // $F2 is column-absolute, row-relative: the rule walks down the rows.
        sheet.ConditionalFormats.Add(sheet.GetRange(1, 0, Rows, 5),
                new FormulaRule("$F2=\"overdue\"")
                {
                    Style = new ConditionalStyle { FillColor = Color.FromRgb(0xFF, 0xEB, 0x9C) },
                });

        Note(sheet, "expression(\"$F2=\\\"overdue\\\"\") — anchored column, relative row, "
                + "so the whole row lights up.");
        return wb;
    }

    private static void Note(Sheet sheet, string text)
    {
        sheet[Rows + 2, 0].Value = text;
        sheet[Rows + 2, 0].Style = new CellStyle { Font = new() { Italic = true } };
    }

    /// <summary>The starting sheet: monthly figures read from a packaged workbook.</summary>
    private static Workbook Figures() => Workbook.Open(DemoBooks.Read(Content), WorkbookFormat.Xlsx);
}
