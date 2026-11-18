using Axyra.Sheets.Content;
using Axyra.Sheets.Style;

namespace Axyra.Sheets.Examples.FeatureDemo;

/// <summary>
/// Data validation — constrain what a cell will accept.
/// </summary>
/// <remarks>
/// Each sample writes a rule over a range and shades that range, so the governed area is
/// legible while it is still empty. <see cref="Validation"/> covers list, whole number,
/// decimal, date, time, text length and arbitrary formula rules. A list rule shows its
/// in-cell arrow unless <c>ShowDropdown = false</c> asks otherwise.
/// </remarks>
public sealed class DataValidationExample : IExample
{
    public string Name => "feature-demo/data-validation";

    public string Description => "Add dropdown, numeric, date, prompt and custom-formula validation rules";

    public void Run(string outputDirectory)
    {
        using (Workbook wb = Dropdown())
        {
            wb.Save(Path.Combine(outputDirectory, "data-validation-dropdown.xlsx"));
        }
        using (Workbook wb = Numeric())
        {
            wb.Save(Path.Combine(outputDirectory, "data-validation-numeric.xlsx"));
        }
        using (Workbook wb = Prompts())
        {
            wb.Save(Path.Combine(outputDirectory, "data-validation-prompts.xlsx"));
        }
        using (Workbook wb = Custom())
        {
            wb.Save(Path.Combine(outputDirectory, "data-validation-custom.xlsx"));
        }
        Console.WriteLine("Wrote four data-validation workbooks to " + outputDirectory);
    }

    /// <summary>An inline list, and one backed by a range of allowed values.</summary>
    private static Workbook Dropdown()
    {
        Workbook wb = new Workbook();
        Sheet sheet = Form(wb, "Region", "Stage");

        sheet.Validations.Add("A2:A20", Validation.List("North", "South", "East", "West"));

        // The allowed values live on the sheet, so they are editable data.
        sheet[0, 4].Value = "Allowed stages";
        sheet["E2:E5"].SetValues(new CellValue[,] { { "Draft" }, { "Review" }, { "Final" }, { "Void" } });
        sheet.Validations.Add("B2:B20", new Validation(ValidationType.List) { Formula1 = "$E$2:$E$5" });

        MarkInputs(sheet, "A2:A20", "B2:B20");
        Hint(sheet, "Columns A and B each carry a dropdown; A lists its values inline, "
                + "B reads them from E2:E5.");
        return wb;
    }

    /// <summary>Ranges, comparisons and dates, all as format-free criteria strings.</summary>
    private static Workbook Numeric()
    {
        Workbook wb = new Workbook();
        Sheet sheet = Form(wb, "Quantity", "Delivery date");

        sheet.Validations.Add("A2:A20",
                Validation.WholeNumber(ValidationOperator.Between, "1", "100"));
        sheet.Validations.Add("B2:B20",
                Validation.Date(ValidationOperator.Between, "DATE(2026,1,1)", "DATE(2026,12,31)"));
        sheet.Validations.Add("C2:C20",
                Validation.Decimal(ValidationOperator.GreaterThanOrEqual, "0", null));
        sheet.Validations.Add("D2:D20",
                Validation.TextLength(ValidationOperator.LessThanOrEqual, "40", null));

        sheet[0, 2].Value = "Unit price";
        sheet[0, 3].Value = "Comment";
        sheet["A1:D1"].Style = new CellStyle { Font = new() { Bold = true } };

        // One row that satisfies every rule, as an example of what each column accepts.
        sheet[1, 0].Value = 25;
        sheet[1, 1].DateValue = new DateOnly(2026, 6, 15);
        sheet[1, 2].Value = 19.99;
        sheet[1, 3].Value = "Ship before month end";

        MarkInputs(sheet, "A2:A20", "B2:B20", "C2:C20", "D2:D20");
        sheet["B2:B20"].Style = InputStyle() with { NumberFormat = "yyyy-mm-dd" };
        sheet["C2:C20"].Style = InputStyle() with { NumberFormat = "#,##0.00" };

        Hint(sheet, "Quantity 1-100, delivery inside 2026, price >= 0, comment <= 40 chars. "
                + "Row 2 holds a value each rule accepts.");
        return wb;
    }

    /// <summary>The message a user sees before typing, and the one that stops them.</summary>
    private static Workbook Prompts()
    {
        Workbook wb = new Workbook();
        Sheet sheet = Form(wb, "Stage", "Headcount");

        sheet["B2:B5"].SetValues(new CellValue[,] { { "Draft" }, { "Review" }, { "Final" }, { "Void" } });
        sheet.Columns[1].Hidden = true;

        sheet.Validations.Add("A2:A20", new Validation(ValidationType.List)
        {
            Formula1 = "$B$2:$B$5",
            AllowBlank = false,
            InputTitle = "Workflow stage",
            InputMessage = "Pick a stage from the list.",
            ErrorTitle = "Not a stage",
            ErrorMessage = "Use the dropdown — free text is rejected.",
            ErrorStyle = ValidationErrorStyle.Stop,
        });

        sheet[1, 0].Value = "Draft";
        MarkInputs(sheet, "A2:A20");

        Hint(sheet, "\"stop\" blocks the entry. \"warning\" and \"information\" let it through.");
        return wb;
    }

    /// <summary>Anything expressible as a formula can be a rule.</summary>
    private static Workbook Custom()
    {
        Workbook wb = new Workbook();
        Sheet sheet = Form(wb, "Multiple of 5", "Order ref");

        sheet.Validations.Add("A2:A20", Validation.Custom("AND(ISNUMBER(A2), MOD(A2,5)=0)"));
        sheet.Validations.Add("B2:B20", Validation.Custom("LEFT(B2,4)=\"ORD-\""));

        sheet[1, 0].Value = 25;
        sheet[1, 1].Value = "ORD-1234";
        MarkInputs(sheet, "A2:A20", "B2:B20");

        Hint(sheet, "Column A accepts multiples of five; column B needs an ORD- prefix. "
                + "The formula is relative to the range's first cell.");
        return wb;
    }

    private static Sheet Form(Workbook wb, string first, string second)
    {
        Sheet sheet = wb.Sheets.Add("Entry");
        sheet[0, 0].Value = first;
        sheet[0, 1].Value = second;
        sheet["A1:B1"].Style = new CellStyle
        {
            Font = new() { Bold = true },
            Fill = Color.FromRgb(0xDD, 0xEB, 0xF7),
        };
        sheet.Columns[0].Width = 20;
        sheet.Columns[1].Width = 20;
        sheet.Columns[2].Width = 16;
        sheet.Columns[3].Width = 20;
        sheet.Columns[4].Width = 18;
        return sheet;
    }

    /// <summary>
    /// The input-cell marking: a <c>#FFF2CC</c> fill inside a <c>#D9A441</c> edge.
    /// Returned as a record so a column can add its own number format with <c>with</c>
    /// without losing the marking.
    /// </summary>
    private static CellStyle InputStyle() => new()
    {
        Fill = Color.FromRgb(0xFF, 0xF2, 0xCC),
        Borders = Borders.Around(BorderStyle.Thin, Color.FromRgb(0xD9, 0xA4, 0x41)),
    };

    /// <summary>
    /// Marks the ranges a rule governs. A validation is invisible until it rejects
    /// something, so shading the governed cells makes them legible while empty.
    /// </summary>
    private static void MarkInputs(Sheet sheet, params string[] ranges)
    {
        foreach (string reference in ranges)
        {
            sheet[reference].Style = InputStyle();
        }
    }

    private static void Hint(Sheet sheet, string text)
    {
        sheet[22, 0].Value = text;
        sheet[22, 0].Style = new CellStyle { Font = new() { Italic = true } };
    }
}
