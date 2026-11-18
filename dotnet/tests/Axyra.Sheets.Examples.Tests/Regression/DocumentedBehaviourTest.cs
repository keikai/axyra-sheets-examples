using Axyra.Sheets.Calc;
using Axyra.Sheets.Content;
using Axyra.Sheets.IO;
using Axyra.Sheets.Pivots;
using Axyra.Sheets.Style;
using static Axyra.Sheets.Examples.Tests.Regression.IssueTestSupport;

namespace Axyra.Sheets.Examples.Tests.Regression;

/// <summary>
/// Behaviours the Axyra documentation and the axyra-sheets skill state, checked against the build.
/// Code written from the documentation relies on each of them; a failure means the text or the
/// engine needs to change.
/// </summary>
public class DocumentedBehaviourTest
{
    /// <summary><c>new Workbook()</c> has no sheets; <c>UsedRange</c> of an empty sheet is null</summary>
    [Fact]
    public void EmptyWorkbookAndSheet()
    {
        using Workbook wb = new();
        Assert.Empty(wb.Sheets);
        Assert.Null(wb.Sheets.Add("Empty").UsedRange);
    }

    /// <summary><c>Validation.List</c> refuses a value that contains a comma</summary>
    [Fact]
    public void ValidationListRejectsCommas()
    {
        Assert.ThrowsAny<ArgumentException>(() => Validation.List("a,b", "c"));
    }

    /// <summary><c>ShowDropdown = false</c> writes showDropDown="1", which hides the arrow; <c>Validation.List</c> does not</summary>
    [Fact]
    public void ShowDropdownFalseHidesTheArrow()
    {
        using TempDir dir = new();
        string p = dir.Resolve("dropdown.xlsx");
        using (Workbook wb = new())
        {
            Sheet s = wb.Sheets.Add("S");
            s.Validations.Add("A1:A3", new Validation(ValidationType.List) { Formula1 = "\"x,y\"", ShowDropdown = false });
            s.Validations.Add("B1:B3", Validation.List("x", "y"));
            wb.Save(p);
        }
        List<string> rules = FindAll(SheetsXml(p), "<dataValidation [^>]*>");
        string all = string.Join(", ", rules);
        Assert.True(rules.Any(r => r.Contains("A1:A3") && r.Contains("showDropDown=\"1\"")), all);
        Assert.True(rules.Any(r => r.Contains("B1:B3") && !r.Contains("showDropDown")), all);
    }

    /// <summary><c>Formula</c> is null for a text value that starts with '='</summary>
    [Fact]
    public void TextStartingWithEqualsStaysText()
    {
        using Workbook wb = new();
        Sheet s = wb.Sheets.Add("S");
        s[0, 0].Value = T("=x");
        Assert.Null(s[0, 0].Formula);
        Assert.Equal(T("=x"), s[0, 0].Value);
    }

    /// <summary>Sheet handles are new objects: <c>Sheets.IndexOf(sheet)</c> is -1, use <c>sheet.Index</c></summary>
    [Fact]
    public void SheetHandlesHaveNoEquality()
    {
        using Workbook wb = new();
        Sheet s = wb.Sheets.Add("S");
        // Java: wb.sheets().indexOf(s); SheetCollection.IndexOf takes a name in .NET, so the list
        // search goes through IReadOnlyList<Sheet>, as Java's List.indexOf does.
        Assert.Equal(-1, wb.Sheets.ToList().IndexOf(s));
        Assert.Equal(0, s.Index);
    }

    // Java: a bean with getZeta() and getAlpha(), read as the properties "zeta" and "alpha".
    // The .NET counterpart is a class with public properties, whose names are PascalCase.
    public class Bean
    {
        public string Zeta => "z";

        public string Alpha => "a";
    }

    /// <summary>ImportData: class properties are ordered by name, dictionary keys in first-met order, a DateOnly without a date format</summary>
    [Fact]
    public void ImportDataColumnOrder()
    {
        using Workbook wb = new();
        Sheet s = wb.Sheets.Add("S");
        s.ImportData(0, 0, new List<Bean> { new() });
        // Java: "alpha", "zeta" (bean property names); .NET: the property names themselves
        Assert.Equal(T("Alpha"), s[0, 0].Value);
        Assert.Equal(T("Zeta"), s[0, 1].Value);

        // Dictionary<,> enumerates in insertion order when nothing is removed, like Java's LinkedHashMap
        Dictionary<string, object> first = new() { ["b"] = 1, ["a"] = 2 };
        Dictionary<string, object> second = new() { ["c"] = 3, ["a"] = 4 };
        List<Dictionary<string, object>> rows = [first, second];
        s.ImportData(3, 0, rows);
        Assert.Equal([T("b"), T("a"), T("c")], new[] { s[3, 0].Value, s[3, 1].Value, s[3, 2].Value });

        s.ImportData(6, 0, new List<Dictionary<string, object>> { new() { ["d"] = new DateOnly(2024, 3, 15) } });
        Assert.Equal(45366.0, Number(s[7, 0].Value));
        Assert.Equal("General", s[7, 0].Style.NumberFormat);
    }

    /// <summary><c>CellRange.Style</c> returns the top-left cell's style</summary>
    [Fact]
    public void RangeStyleIsTopLeft()
    {
        using Workbook wb = new();
        Sheet s = wb.Sheets.Add("S");
        s[0, 0].Style = new CellStyle { Font = new Font { Bold = true } };
        Assert.True(s["A1:B1"].Style.Font.Bold);
    }

    /// <summary><c>Color.ToHex()</c> throws for a theme colour</summary>
    [Fact]
    public void ToHexNeedsAnRgbColour()
    {
        Assert.False(Color.FromTheme(1).IsRgb);
        // Java: IllegalStateException; .NET: InvalidOperationException (16-design-conventions, Errors)
        Assert.Throws<InvalidOperationException>(() => Color.FromTheme(1).ToHex());
    }

    /// <summary><c>Workbook.Theme</c> is null on a new workbook</summary>
    [Fact]
    public void NewWorkbookHasNoTheme()
    {
        using Workbook wb = new();
        Assert.Null(wb.Theme);
    }

    /// <summary>an API-created pivot is blank in memory until <c>Refresh()</c></summary>
    [Fact]
    public void PivotIsBlankUntilRefresh()
    {
        using Workbook wb = new();
        Sheet data = wb.Sheets.Add("Data");
        data[0, 0].Value = T("R");
        data[0, 1].Value = T("V");
        data[1, 0].Value = T("N");
        data[1, 1].Value = N(5);
        Sheet pivot = wb.Sheets.Add("Pv");
        PivotTable pv = pivot.PivotTables.Add("P", "A1", "Data!$A$1:$B$2");
        pv.RowFields.Add("R");
        pv.DataFields.Add("V", AggregateFunction.Sum);
        Assert.Null(pivot.UsedRange);
        pv.Refresh();
        Assert.Equal(T("Row Labels"), pivot[0, 0].Value);
    }

    /// <summary><c>Subtotal</c> returns the number of rows it inserted</summary>
    [Fact]
    public void SubtotalReturnsInsertedRows()
    {
        using Workbook wb = new();
        Sheet s = wb.Sheets.Add("S");
        string[][] rows = [["Cat", "Amt"], ["a", "1"], ["a", "2"], ["b", "3"]];
        for (int r = 0; r < rows.Length; r++)
        {
            s[r, 0].Value = T(rows[r][0]);
            s[r, 1].Value = r == 0 ? T(rows[r][1]) : N(double.Parse(rows[r][1]));
        }
        // Java: s.subtotal(firstRow 0, lastRow 3, groupBy 0, hasHeader true, SUM, column 1)
        Assert.Equal(3, s["A1:B4"].Subtotal(0, SubtotalFunction.Sum, [1], hasHeaders: true));
    }

    /// <summary><c>SavePdf</c> throws for an unknown sheet name</summary>
    [Fact]
    public void UnknownPdfSheetThrows()
    {
        using TempDir dir = new();
        using Workbook wb = new();
        wb.Sheets.Add("S");
        Assert.ThrowsAny<Exception>(
            () => wb.SavePdf(dir.Resolve("x.pdf"), new PdfOptions { SheetNames = ["NoSuch"] }));
    }

    /// <summary><c>CalcMode.Manual</c> is saved to and read from XLSX</summary>
    [Fact]
    public void ManualModeRoundTrips()
    {
        using TempDir dir = new();
        string p = dir.Resolve("manual.xlsx");
        using (Workbook wb = new())
        {
            wb.Sheets.Add("S")[0, 0].Value = N(1);
            wb.CalcMode = CalcMode.Manual;
            wb.Save(p);
        }
        using (Workbook wb = Workbook.Open(p))
        {
            Assert.Equal(CalcMode.Manual, wb.CalcMode);
        }
    }

    /// <summary>ODS writes a formula without an OpenFormula translation as its value; Keep makes that save throw</summary>
    [Fact]
    public void OdsFormulaPolicy()
    {
        using TempDir dir = new();
        using (Workbook wb = new())
        {
            Sheet s = wb.Sheets.Add("S");
            s[0, 0].Formula = "XLOOKUP(1,{1,2},{3,4})";
            s[1, 0].Formula = "SUM(1,2)";
            wb.Save(dir.Resolve("plain.ods"));
            ArgumentException e = Assert.ThrowsAny<ArgumentException>(() => wb.Save(
                dir.Resolve("keep.ods"), new SaveOptions { FormulaPolicy = FormulaPolicy.Keep }));
            Assert.True(e.Message.Contains("S!A1"), e.Message);
        }
        using (Workbook wb = Workbook.Open(dir.Resolve("plain.ods")))
        {
            Sheet s = wb.Sheets["S"];
            Assert.Null(s[0, 0].Formula);
            Assert.Equal(3.0, Number(s[0, 0].Value));
            Assert.Equal("SUM(1,2)", s[1, 0].Formula);
        }
    }

    /// <summary>CSV defaults to formatted text with CRLF line ends</summary>
    [Fact]
    public void CsvDefaults()
    {
        using TempDir dir = new();
        string p = dir.Resolve("formatted.csv");
        using (Workbook wb = new())
        {
            Sheet s = wb.Sheets.Add("S");
            s[0, 0].Value = N(1234.5);
            s[0, 0].Style = new CellStyle { NumberFormat = "#,##0.00" };
            wb.SaveCsv(p, new CsvOptions { SheetName = "S" });
            Assert.Contains("\r\n", System.Text.Encoding.UTF8.GetString(wb.ToByteArray(WorkbookFormat.Csv)));
        }
        Assert.True(File.ReadAllText(p).StartsWith("\"1,234.50\"", StringComparison.Ordinal), File.ReadAllText(p));
    }

    /// <summary><c>StreamWorkbook</c> refuses rows out of order and drops an uncommitted row</summary>
    [Fact]
    public void StreamRowContract()
    {
        using TempDir dir = new();
        Assert.ThrowsAny<Exception>(() =>
        {
            using StreamWorkbook sw = new(dir.Resolve("order.xlsx"));
            sw.StartSheet("A");
            sw.StartRow(2).AddNumber(1).Commit();
            sw.StartRow(1).AddNumber(2).Commit();
        });
        string p = dir.Resolve("uncommitted.xlsx");
        using (StreamWorkbook sw = new(p))
        {
            sw.StartSheet("A");
            sw.StartRow(0).AddNumber(1).Commit();
            sw.StartRow(1).AddNumber(2);
            sw.StartRow(2).AddNumber(3).Commit();
            sw.Finish();
        }
        using (Workbook wb = Workbook.Open(p))
        {
            Assert.IsType<CellValue.Blank>(wb.Sheets["A"][1, 0].Value);
            Assert.Equal(3.0, Number(wb.Sheets["A"][2, 0].Value));
        }
    }
}
