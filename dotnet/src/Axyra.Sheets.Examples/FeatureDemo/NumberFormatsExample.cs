using System.Globalization;
using Axyra.Sheets.Format;
using Axyra.Sheets.Style;

namespace Axyra.Sheets.Examples.FeatureDemo;

/// <summary>Number and date formats — Excel format codes, applied and evaluated.</summary>
/// <remarks>
/// A format never changes the stored value, only its presentation: <c>FormattedText</c> returns
/// what a viewer shows, <c>Value</c> the number underneath. <see cref="Formats"/> does the same
/// job statelessly, with no workbook involved, and both are locale-aware.
/// </remarks>
public sealed class NumberFormatsExample : IExample
{
    public string Name => "feature-demo/number-formats";

    public string Description => "Apply number and date format codes, per locale and without a workbook";

    public void Run(string outputDirectory)
    {
        using (Workbook wb = Numbers())
        {
            wb.Save(Path.Combine(outputDirectory, "number-formats-numbers.xlsx"));
        }
        using (Workbook wb = Dates())
        {
            wb.Save(Path.Combine(outputDirectory, "number-formats-dates.xlsx"));
        }
        using (Workbook wb = Locales())
        {
            wb.Save(Path.Combine(outputDirectory, "number-formats-locales.xlsx"));
        }
        using (Workbook wb = Stateless())
        {
            wb.Save(Path.Combine(outputDirectory, "number-formats-stateless.xlsx"));
        }
        Console.WriteLine("Wrote four number-format workbooks to " + outputDirectory);
    }

    private static Workbook Numbers()
    {
        Workbook wb = new Workbook();
        Sheet sheet = Header(wb, "Numbers", "Format code", "Formatted", "Underlying value");

        string[] codes =
        [
            "General", "#,##0", "#,##0.00", "0.0%", "0.0%;[Red]-0.0%",
            "$#,##0.00", "[$€-407]#,##0.00", "0.00E+00", "# ?/?", "\"OK\";\"OK\";\"zero\"",
        ];
        double[] values = [1234.5678, 1234.5678, 1234.5678, 0.0742, -0.0742,
            1234.5678, 1234.5678, 1234.5678, 1.75, 1];

        for (int i = 0; i < codes.Length; i++)
        {
            int row = i + 1;
            sheet[row, 0].Value = codes[i];
            sheet[row, 1].Value = values[i];
            sheet.GetRange(row, 1, row, 1).Style = new CellStyle { NumberFormat = codes[i] };
            sheet[row, 2].Value = values[i].ToString(CultureInfo.InvariantCulture);
        }
        return wb;
    }

    private static Workbook Dates()
    {
        Workbook wb = new Workbook();
        Sheet sheet = Header(wb, "Dates", "Format code", "Formatted", "Serial");

        string[] codes =
        [
            "yyyy-mm-dd", "d mmm yyyy", "dddd d mmmm yyyy",
            "mm/dd/yy", "mmm-yy", "hh:mm:ss", "yyyy-mm-dd hh:mm",
        ];
        DateOnly date = new DateOnly(2026, 8, 24);

        for (int i = 0; i < codes.Length; i++)
        {
            int row = i + 1;
            sheet[row, 0].Value = codes[i];
            sheet[row, 1].DateValue = date;
            sheet.GetRange(row, 1, row, 1).Style = new CellStyle { NumberFormat = codes[i] };
            sheet[row, 2].Value = sheet[row, 1].Value;
        }
        sheet[codes.Length + 2, 0].Value = "Every row holds the same serial — only the format differs.";
        return wb;
    }

    /// <summary>Month names, separators and grouping follow the locale.</summary>
    private static Workbook Locales()
    {
        Workbook wb = new Workbook();
        Sheet sheet = Header(wb, "Locales", "Locale", "dddd d mmmm yyyy", "#,##0.00");

        string[] tags = ["en-US", "de-DE", "fr-FR", "ja-JP"];
        for (int i = 0; i < tags.Length; i++)
        {
            int row = i + 1;
            sheet[row, 0].Value = tags[i];
            sheet[row, 1].Value = Formats.Format(46_258, "dddd d mmmm yyyy", tags[i]);
            sheet[row, 2].Value = Formats.Format(1234.5678, "#,##0.00", tags[i]);
        }
        sheet.Columns[1].Width = 30;
        return wb;
    }

    /// <summary><see cref="Formats"/> formats and parses without a workbook at all.</summary>
    private static Workbook Stateless()
    {
        Workbook wb = new Workbook();
        Sheet sheet = Header(wb, "Formats", "Call", "Result", "");

        Row(sheet, 1, "Formats.Format(1234.5678, \"$#,##0.00\")",
            Formats.Format(1234.5678, "$#,##0.00"));
        Row(sheet, 2, "Formats.Format(0.0742, \"0.0%\")",
            Formats.Format(0.0742, "0.0%"));
        Row(sheet, 3, "Formats.Format(46258, \"d mmm yyyy\")",
            Formats.Format(46_258, "d mmm yyyy"));
        Row(sheet, 4, "Formats.Parse(\"1,234.56\")",
            Describe(Formats.Parse("1,234.56")));
        Row(sheet, 5, "Formats.Parse(\"1 234,56\")",
            Describe(Formats.Parse("1 234,56")));
        sheet[7, 0].Value = "Parse reports what it recognised. A string it cannot read as a number "
            + "comes back as Text rather than as an error.";

        sheet.Columns[0].Width = 42;
        sheet.Columns[1].Width = 24;
        return wb;
    }

    /// <summary>A parse result rendered as what it means.</summary>
    /// <remarks>
    /// <see cref="ParsedValue"/> is a tagged union over one record: <c>Number</c> is the payload
    /// for Number, Date and Bool, and <c>Text</c> is the payload for Text and Formula, so only
    /// the field that carries the value is shown.
    /// </remarks>
    private static string Describe(ParsedValue parsed) => parsed.Kind switch
    {
        ParsedValue.InputKind.Number or ParsedValue.InputKind.Date =>
            parsed.Kind + " " + parsed.Number.ToString(CultureInfo.InvariantCulture),
        ParsedValue.InputKind.Bool => parsed.Kind + " " + parsed.BooleanValue,
        _ => parsed.Kind + " \"" + parsed.Text + "\"",
    };

    private static void Row(Sheet sheet, int index, string call, string result)
    {
        sheet[index, 0].Value = call;
        sheet[index, 1].Value = result;
    }

    private static Sheet Header(Workbook wb, string name, string a, string b, string c)
    {
        Sheet sheet = wb.Sheets.Add(name);
        sheet["A1:C1"].SetValues(new CellValue[,] { { a, b, c } });
        sheet["A1:C1"].Style = new CellStyle { Font = new Font { Bold = true } };
        sheet.Columns[0].Width = 26;
        sheet.Columns[1].Width = 24;
        sheet.Columns[2].Width = 20;
        return sheet;
    }
}
