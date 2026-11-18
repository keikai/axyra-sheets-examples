using Axyra.Sheets.Content;
using Axyra.Sheets.Format;
using Axyra.Sheets.Style;
using static Axyra.Sheets.Examples.Tests.Regression.IssueTestSupport;

namespace Axyra.Sheets.Examples.Tests.Regression;

/// <summary>
/// Features F01-F13 of the axyra-sheets skill test (workbooks, cell values, formulas, names, styles,
/// rich text, comments, hyperlinks). Each test does what the feature task asked and checks the result
/// the way the test judge did: in memory and, where the task saves, after reopening the file.
/// </summary>
public class FeatureCoreTest
{
    /// <summary>F01 create a workbook with two sheets and typed values, save, reopen</summary>
    [Fact]
    public void CreateAndSave()
    {
        using TempDir dir = new();
        string p = dir.Resolve("data.xlsx");
        using (Workbook wb = new())
        {
            Sheet data = wb.Sheets.Add("Data");
            wb.Sheets.Add("Summary");
            data["A1:C1"].SetValues(new CellValue[,] { { T("Hello"), N(42.5), CellValue.Of(true) } });
            wb.Save(p);
        }
        using (Workbook wb = Workbook.Open(p))
        {
            Sheet data = wb.Sheets["Data"];
            Assert.NotNull(wb.Sheets["Summary"]);
            Assert.Equal(T("Hello"), data[0, 0].Value);
            Assert.Equal(42.5, Number(data[0, 1].Value));
            Assert.Equal(CellValue.Of(true), data[0, 2].Value);
        }
    }

    /// <summary>F02 open a saved workbook and read every row</summary>
    [Fact]
    public void OpenAndRead()
    {
        using TempDir dir = new();
        string p = dir.Resolve("input.xlsx");
        using (Workbook wb = new())
        {
            wb.Sheets.Add("Data").GetRange(0, 0, 3, 1).SetValues(new CellValue[,]
            {
                { T("name"), T("qty") }, { T("Apple"), N(10) }, { T("Banana"), N(25) }, { T("Cherry"), N(7) },
            });
            wb.Save(p);
        }
        using (Workbook wb = Workbook.Open(p))
        {
            CellValue[,] rows = wb.Sheets["Data"].GetRange(0, 0, 3, 1).GetValues();
            Assert.Equal(T("qty"), rows[0, 1]);
            Assert.Equal(T("Cherry"), rows[3, 0]);
            Assert.Equal(7.0, Number(rows[3, 1]));
        }
    }

    /// <summary>F03 typed values, an error formula and formatted text</summary>
    [Fact]
    public void TypedValues()
    {
        using Workbook wb = new();
        Sheet s = wb.Sheets.Add("S");
        s[0, 0].Value = T("Hello");
        s[1, 0].Value = N(1234.5678);
        s[2, 0].Value = CellValue.Of(true);
        s[4, 0].Formula = "1/0";
        s[5, 0].Value = N(1234567.891);
        s[5, 0].Style = new CellStyle { NumberFormat = "#,##0.00" };
        Assert.IsType<CellValue.Text>(s[0, 0].Value);
        Assert.IsType<CellValue.Number>(s[1, 0].Value);
        Assert.IsType<CellValue.Bool>(s[2, 0].Value);
        Assert.IsType<CellValue.Blank>(s[3, 0].Value);
        Assert.IsType<CellValue.Error>(s[4, 0].Value);
        Assert.Equal("1,234,567.89", s[5, 0].FormattedText);
    }

    /// <summary>F04 dates and date-times read back as DateOnly and DateTime (Java: LocalDate and LocalDateTime)</summary>
    [Fact]
    public void DatesAndTimes()
    {
        using Workbook wb = new();
        Sheet s = wb.Sheets.Add("S");
        s[0, 0].DateValue = new DateOnly(2024, 3, 15);
        s[1, 0].DateTimeValue = new DateTime(2024, 3, 15, 14, 30, 0);
        Assert.Equal(new DateOnly(2024, 3, 15), s[0, 0].DateValue);
        Assert.Equal(new DateTime(2024, 3, 15, 14, 30, 0), s[1, 0].DateTimeValue);
    }

    /// <summary>F05 bulk write and read a 100 x 3 block</summary>
    [Fact]
    public void BulkReadWrite()
    {
        using Workbook wb = new();
        Sheet s = wb.Sheets.Add("S");
        CellValue[,] data = new CellValue[100, 3];
        for (int r = 0; r < 100; r++)
            for (int c = 0; c < 3; c++)
                data[r, c] = N(r + c);
        s.GetRange(0, 0, 99, 2).SetValues(data);
        CellValue[,] read = s.GetRange(0, 0, 99, 2).GetValues();
        double sum = 0;
        for (int r = 0; r < read.GetLength(0); r++)
            sum += Number(read[r, 1]);
        Assert.Equal(100, read.GetLength(0));
        Assert.Equal(5050.0, sum);
    }

    /// <summary>F06 formulas recalculate after an input changes; evaluate an ad-hoc formula</summary>
    [Fact]
    public void FormulasAndRecalculation()
    {
        using Workbook wb = new();
        Sheet s = wb.Sheets.Add("S");
        for (int r = 0; r < 5; r++)
            s[r, 0].Value = N((r + 1) * 10);
        s[5, 0].Formula = "SUM(A1:A5)";
        s[6, 0].Formula = "AVERAGE(A1:A5)";
        wb.Recalculate();
        Assert.Equal("SUM(A1:A5)", s[5, 0].Formula);
        Assert.Equal(150.0, Number(s[5, 0].Value));
        Assert.Equal(30.0, Number(s[6, 0].Value));
        s[0, 0].Value = N(100);
        wb.Recalculate();
        Assert.Equal(240.0, Number(s[5, 0].Value));
        Assert.Equal(200.0, Number(s.Evaluate("MAX(A1:A5)*2")));
        Assert.IsType<CellValue.Blank>(s[7, 0].Value);
    }

    /// <summary>F07 CSE array formula and spilling dynamic arrays, also after reopening</summary>
    [Fact]
    public void ArrayFormulas()
    {
        using TempDir dir = new();
        string p = dir.Resolve("arrays.xlsx");
        using (Workbook wb = new())
        {
            Sheet s = wb.Sheets.Add("S");
            double[,] ab = { { 1, 10 }, { 5, 3 }, { 3, 2.5 }, { 4, 1.25 }, { 1, 7 } };
            for (int r = 0; r < ab.GetLength(0); r++)
            {
                s[r + 1, 0].Value = N(ab[r, 0]);
                s[r + 1, 1].Value = N(ab[r, 1]);
            }
            s["C2:C6"].SetArrayFormula("A2:A6*B2:B6");
            s[1, 4].Formula = "SEQUENCE(5,1,10,10)";
            Assert.Equal(10.0, Number(s[1, 2].Value));
            Assert.Equal(15.0, Number(s[2, 2].Value));
            Assert.Equal("E2:E6", s["E2"].SpillRange());
            Assert.Equal(50.0, Number(s[5, 4].Value));
            wb.Save(p);
        }
        using (Workbook wb = Workbook.Open(p))
        {
            Sheet s = wb.Sheets["S"];
            Assert.Equal("E2:E6", s["E2"].SpillRange());
            Assert.Equal(7.0, Number(s[5, 2].Value));
        }
    }

    /// <summary>F08 a workbook-level name used in a formula, listed and removed</summary>
    [Fact]
    public void NamedRanges()
    {
        using Workbook wb = new();
        Sheet s = wb.Sheets.Add("Sheet1");
        for (int r = 1; r <= 4; r++)
            s[r, 1].Value = N(r * 75);
        wb.Names.Add("Sales", "Sheet1!$B$2:$B$5");
        s[6, 1].Formula = "SUM(Sales)";
        wb.Recalculate();
        Assert.Equal(750.0, Number(s[6, 1].Value));
        Assert.Equal(["Sales"], wb.Names.Select(x => x.Name).ToList());
        Assert.True(wb.Names.Remove("Sales"));
        // Java: findName returns null; .NET: TryGet returns false (16-design-conventions, Collections)
        Assert.False(wb.Names.TryGet("Sales", out _));
    }

    /// <summary>F09 header and body cell styles survive save and reopen</summary>
    [Fact]
    public void CellStyling()
    {
        using TempDir dir = new();
        string p = dir.Resolve("styled.xlsx");
        using (Workbook wb = new())
        {
            Sheet s = wb.Sheets.Add("S");
            s[0, 0].Value = T("Header");
            s["A1:C1"].Style = new CellStyle
            {
                Font = new Font { Bold = true, Color = Color.FromRgb(255, 255, 255) },
                Fill = Fill.Solid(Color.FromRgb(0x1F, 0x3A, 0x5F)),
                Borders = new Borders { Bottom = new BorderEdge(BorderStyle.Thin, Color.FromRgb(0, 0, 0)) },
                HorizontalAlignment = HorizontalAlignment.Center,
            };
            s[1, 2].Value = T("Body text that wraps");
            s[1, 2].Style = new CellStyle
            {
                Font = new Font { Italic = true, Color = Color.FromRgb(255, 0, 0) },
                WrapText = true,
            };
            wb.Save(p);
        }
        using (Workbook wb = Workbook.Open(p))
        {
            CellStyle header = wb.Sheets["S"][0, 0].Style;
            CellStyle body = wb.Sheets["S"][1, 2].Style;
            Assert.True(header.Font.Bold);
            Assert.Equal("FF1F3A5F", header.Fill.Foreground!.ToHex().Replace("#", ""));
            Assert.Equal(BorderStyle.Thin, header.Borders.Bottom.Style);
            Assert.Equal(HorizontalAlignment.Center, header.HorizontalAlignment);
            Assert.True(body.Font.Italic);
            Assert.True(body.WrapText);
            Assert.Equal("FFFF0000", body.Font.Color!.ToHex().Replace("#", ""));
        }
    }

    /// <summary>F10 number formats on cells and the standalone format engine</summary>
    [Fact]
    public void NumberFormats()
    {
        using (Workbook wb = new())
        {
            Sheet s = wb.Sheets.Add("S");
            Cell currency = s[0, 0];
            currency.Value = N(1234.5);
            currency.Style = new CellStyle { NumberFormat = "$#,##0.00" };
            Cell percent = s[1, 0];
            percent.Value = N(0.1234);
            percent.Style = new CellStyle { NumberFormat = "0.00%" };
            Cell thousands = s[2, 0];
            thousands.Value = N(9876543.21);
            thousands.Style = new CellStyle { NumberFormat = "#,##0" };
            Cell date = s[3, 0];
            date.DateValue = new DateOnly(2026, 3, 15);
            date.Style = date.Style with { NumberFormat = "yyyy-mm-dd ddd" };
            Assert.Equal("$1,234.50", currency.FormattedText);
            Assert.Equal("12.34%", percent.FormattedText);
            Assert.Equal("9,876,543", thousands.FormattedText);
            Assert.Equal("2026-03-15 Sun", date.FormattedText);
        }
        Assert.Equal("$1,234.50", Formats.Format(1234.5, "$#,##0.00"));
        Assert.Equal(0.125, Formats.Parse("12.5%").Number);
    }

    /// <summary>F11 rich text runs survive save and reopen</summary>
    [Fact]
    public void RichText()
    {
        using TempDir dir = new();
        string p = dir.Resolve("rich.xlsx");
        using (Workbook wb = new())
        {
            wb.Sheets.Add("S")[0, 0].RichText = new RichText(
                new TextRun("Important") { Bold = true, Color = Color.FromRgb(255, 0, 0) },
                new TextRun(": please review"));
            wb.Save(p);
        }
        using (Workbook wb = Workbook.Open(p))
        {
            IReadOnlyList<TextRun> runs = wb.Sheets["S"][0, 0].RichText!.Runs;
            Assert.Equal(2, runs.Count);
            Assert.Equal("Important", runs[0].Text);
            Assert.True(runs[0].Bold);
            Assert.False(runs[1].Bold);
        }
    }

    /// <summary>F12 a note and a threaded comment with a reply survive save and reopen</summary>
    [Fact]
    public void Comments()
    {
        using TempDir dir = new();
        string p = dir.Resolve("comments.xlsx");
        using (Workbook wb = new())
        {
            Sheet s = wb.Sheets.Add("S");
            s[0, 1].Value = N(1);
            s[0, 1].Comment = new Comment("Alice", "Q3 revenue, unaudited.");
            s[1, 1].Value = N(2);
            ThreadedComment thread = s[1, 1].ThreadedComments.Add("Bob", "Is this final?");
            thread.Reply("Carol", "Yes, confirmed.");
            wb.Save(p);
        }
        using (Workbook wb = Workbook.Open(p))
        {
            Sheet s = wb.Sheets["S"];
            Assert.Equal("Alice", s[0, 1].Comment!.Author);
            Assert.Contains("Q3 revenue", s[0, 1].Comment!.Text);
            ThreadedCommentCollection thread = s[1, 1].ThreadedComments;
            Assert.Equal("Is this final?", thread[0].Text);
            Assert.Equal("Carol", thread[0].Replies[0].Author);
        }
    }

    /// <summary>F13 external and internal hyperlinks survive save and reopen; one is removed</summary>
    [Fact]
    public void Hyperlinks()
    {
        using TempDir dir = new();
        string p = dir.Resolve("links.xlsx");
        using (Workbook wb = new())
        {
            Sheet links = wb.Sheets.Add("Links");
            wb.Sheets.Add("Data")[1, 1].Value = N(1);
            links[0, 0].Value = T("Visit");
            links[0, 0].Hyperlink = Hyperlink.Url("https://www.example.com/");
            links[1, 0].Value = T("Go to Data!B2");
            links[1, 0].Hyperlink = Hyperlink.Internal("Data!B2", "Go to Data!B2");
            wb.Save(p);
        }
        using (Workbook wb = Workbook.Open(p))
        {
            Sheet links = wb.Sheets["Links"];
            Assert.Equal("https://www.example.com/", links[0, 0].Hyperlink!.Address);
            Assert.Contains("Data!B2", links[1, 0].Hyperlink!.Address);
            links[0, 0].Hyperlink = null;
            Assert.Null(links[0, 0].Hyperlink);
            Assert.NotNull(links[1, 0].Hyperlink);
        }
    }
}
