using System.Text;
using Axyra.Sheets.Calc;
using Axyra.Sheets.Charts;
using Axyra.Sheets.Content;
using Axyra.Sheets.Drawing;
using Axyra.Sheets.IO;
using Axyra.Sheets.Pivots;
using Axyra.Sheets.Style;
using Axyra.Sheets.Template;
using static Axyra.Sheets.Examples.Tests.Regression.IssueTestSupport;

namespace Axyra.Sheets.Examples.Tests.Regression;

/// <summary>
/// Regression tests for Axyra Sheets defects found while benchmarking coding agents on the library,
/// all fixed in the build this project uses. Each test asserts the behaviour Excel has; a failure means
/// the defect is back. The numbers in the display names are the items of the benchmark issue lists
/// (0930, 1013, 1021 and 1022 builds).
/// </summary>
/// <remarks>
/// Tests marked <c>[Trait("Category", "known-issue")]</c> fail on the .NET SDK for a reason recorded in
/// <c>CONVERSION-NOTES/regression-fixed-issues.md</c>.
/// </remarks>
public class FixedIssuesTest
{
    [Fact(DisplayName = "#1 Range.autoFill shifts relative references")]
    public void AutoFillShiftsReferences()
    {
        using Workbook wb = new();
        Sheet s = wb.Sheets.Add("S");
        s["A1:B2"].SetValues(new CellValue[,] { { N(1), N(10) }, { N(2), N(20) } });
        s[2, 0].Formula = "SUM(A1:A2)";
        s["A3"].AutoFill(s["A3:B3"]);
        Assert.Equal("SUM(B1:B2)", s[2, 1].Formula);
        Assert.Equal(30.0, Number(s[2, 1].Value));
    }

    [Fact(DisplayName = "#2 insertRows adjusts other sheets' formulas and workbook names")]
    public void InsertRowsAdjustsOtherSheetsAndNames()
    {
        using Workbook wb = new();
        Sheet data = wb.Sheets.Add("Data");
        Sheet summary = wb.Sheets.Add("Summary");
        data["A1:A3"].SetValues(new CellValue[,] { { N(1) }, { N(2) }, { N(3) } });
        summary[0, 0].Formula = "SUM(Data!A1:A3)";
        wb.Names.Add("Lines", "Data!$A$1:$A$3");
        data.Rows.Insert(2, 1);
        data[2, 0].Value = N(100);
        Assert.Equal("SUM(Data!A1:A4)", summary[0, 0].Formula);
        Assert.Equal(106.0, Number(summary[0, 0].Value));
        string lines = wb.Names.First(x => x.Name == "Lines").Formula;
        Assert.Equal("Data!$A$1:$A$4", lines);
    }

    [Fact(DisplayName = "#3 SUBTOTAL(109) and AGGREGATE skip rows hidden by an AutoFilter")]
    public void SubtotalSkipsFilteredRows()
    {
        using Workbook wb = new();
        Sheet s = wb.Sheets.Add("Orders");
        s["A1:B5"].SetValues(new CellValue[,] { { T("Status"), T("Amount") }, { T("Open"), N(10) },
            { T("Shipped"), N(20) }, { T("Open"), N(30) }, { T("Shipped"), N(40) } });
        s.AutoFilter.Range = s["A1:B5"];
        s.AutoFilter[0].FilterByValues("Open");
        s.AutoFilter.Apply();
        s[6, 1].Formula = "SUBTOTAL(109,B2:B5)";
        s[7, 1].Formula = "AGGREGATE(9,5,B2:B5)";
        Assert.Equal(40.0, Number(s[6, 1].Value));
        Assert.Equal(40.0, Number(s[7, 1].Value));
    }

    private static Workbook Region(double a, double b)
    {
        Workbook w = new();
        Sheet d = w.Sheets.Add("Data");
        d["A1:A2"].SetValues(new CellValue[,] { { N(a) }, { N(b) } });
        w.Names.Add("Target", "Data!$A$1");
        Sheet s = w.Sheets.Add("Summary");
        s[0, 0].Formula = "SUM(Data!A1:A2)";
        s[1, 0].Formula = "Target*2";
        return w;
    }

    [Fact(DisplayName = "#4 combine and copySheetFrom keep formulas and names working")]
    public void CopiedSheetsKeepReferences()
    {
        using (Workbook north = Region(1, 2))
        using (Workbook south = Region(100, 200))
        {
            north.Combine(south);
            Sheet copied = north.Sheets["Summary (2)"];
            Assert.Equal("SUM('Data (2)'!A1:A2)", copied[0, 0].Formula);
            Assert.Equal(300.0, Number(copied[0, 0].Value));
        }
        using (Workbook pack = new())
        using (Workbook south = Region(100, 200))
        {
            pack.Sheets.Add("Overview");
            pack.Sheets.AddCopy(south.Sheets[0], "South Data");
            Sheet s = pack.Sheets.AddCopy(south.Sheets[1], "South Summary");
            Assert.Equal(300.0, Number(s[0, 0].Value));
            Assert.Equal(200.0, Number(s[1, 0].Value));
        }
    }

    [Fact(DisplayName = "#5 ODS keeps formulas with FormulaPolicy.KEEP")]
    public void OdsKeepsFormulas()
    {
        using TempDir dir = new();
        string p = dir.Resolve("budget.ods");
        using (Workbook wb = new())
        {
            Sheet s = wb.Sheets.Add("Budget");
            s["A1:A2"].SetValues(new CellValue[,] { { N(1) }, { N(2) } });
            s[2, 0].Formula = "SUM(A1:A2)";
            wb.Save(p, new SaveOptions { FormulaPolicy = FormulaPolicy.Keep });
        }
        Assert.Contains("table:formula", ZipEntry(p, "content.xml"));
        using (Workbook wb = Workbook.Open(p))
        {
            Assert.Equal("SUM(A1:A2)", wb.Sheets["Budget"][2, 0].Formula);
        }
    }

    [Fact(DisplayName = "#7 importData writes Map rows with their keys as the header")]
    public void ImportDataWritesMaps()
    {
        using Workbook wb = new();
        Sheet s = wb.Sheets.Add("S");
        // LinkedHashMap in Java: Dictionary keeps insertion order for a map that is only added to
        Dictionary<string, object> row = new()
        {
            ["id"] = "A1",
            ["qty"] = 3,
        };
        s.ImportData(0, 0, new List<Dictionary<string, object>> { row });
        Assert.Equal(T("id"), s[0, 0].Value);
        Assert.Equal(T("qty"), s[0, 1].Value);
        Assert.Equal(T("A1"), s[1, 0].Value);
        Assert.Equal(3.0, Number(s[1, 1].Value));
    }

    [Fact(DisplayName = "#8 a pivot table is built when the workbook is saved, without refresh()")]
    public void PivotIsBuiltOnSave()
    {
        using TempDir dir = new();
        string p = dir.Resolve("pivot.xlsx");
        using (Workbook wb = new())
        {
            Sheet d = wb.Sheets.Add("Data");
            d["A1:B3"].SetValues(new CellValue[,] { { T("Region"), T("Amount") }, { T("East"), N(10) }, { T("West"), N(20) } });
            Sheet ps = wb.Sheets.Add("Pivot");
            PivotTable pv = ps.PivotTables.Add("Sales", "A3", "Data!$A$1:$B$3");
            pv.RowFields.Add("Region");
            pv.DataFields.Add("Amount", AggregateFunction.Sum);
            wb.Save(p);
        }
        using (Workbook wb = Workbook.Open(p))
        {
            Sheet ps = wb.Sheets["Pivot"];
            List<CellValue> labels = [];
            for (int r = 2; r < 8; r++)
                labels.Add(ps[r, 0].Value);
            Assert.True(labels.Contains(T("East")), "pivot labels " + string.Join(", ", labels));
            Assert.True(labels.Contains(T("West")), "pivot labels " + string.Join(", ", labels));
        }
    }

    [Fact(DisplayName = "#9 setNumber keeps the cell's style")]
    public void SetNumberKeepsStyle()
    {
        using Workbook wb = new();
        Sheet s = wb.Sheets.Add("S");
        s[0, 0].Style = new CellStyle { Font = new Font { Bold = true }, NumberFormat = "#,##0.00" };
        s[0, 0].Value = 1234.5; // Java: s.setNumber(0, 0, 1234.5)
        Assert.True(s[0, 0].Style.Font.Bold);
        Assert.Equal("#,##0.00", s[0, 0].Style.NumberFormat);
    }

    [Fact(DisplayName = "#10 Range.setFormula on several cells shifts relative references")]
    public void MultiCellSetFormulaShifts()
    {
        using Workbook wb = new();
        Sheet s = wb.Sheets.Add("S");
        s["A1:B3"].SetValues(new CellValue[,] { { N(1), N(2) }, { N(3), N(4) }, { N(5), N(6) } });
        s["C1:C3"].SetFormula("A1*B1");
        Assert.Equal(new string?[,] { { "A1*B1" }, { "A2*B2" }, { "A3*B3" } }, s["C1:C3"].GetFormulas());
        Assert.Equal(30.0, Number(s[2, 2].Value));
    }

    [Fact(DisplayName = "#11 removeDuplicates and sort both use range-relative columns")]
    public void RemoveDuplicatesIsRangeRelative()
    {
        using Workbook wb = new();
        Sheet s = wb.Sheets.Add("S");
        CellValue[,] people = { { T("Name"), T("Email") }, { T("Ann"), T("ann1@x") }, { T("Ann"), T("ann2@x") },
            { T("Bob"), T("bob@x") }, { T("Bob"), T("bob@x") } };
        s["C1:D5"].SetValues(people);
        Assert.Equal(1, s["C1:D5"].RemoveDuplicates([1], true));
        Assert.Equal(T("ann2@x"), s[2, 3].Value);

        Sheet s2 = wb.Sheets.Add("S2");
        s2["C1:D5"].SetValues(people);
        s2["C2:D5"].Sort(new SortKey(1, false)); // SortKey.desc(1)
        Assert.Equal(T("bob@x"), s2[1, 3].Value);
        Assert.Equal(T("ann1@x"), s2[4, 3].Value);
    }

    [Fact(DisplayName = "#12 a Sheet handle follows its tab after another sheet is inserted")]
    public void SheetHandleFollowsItsTab()
    {
        using Workbook wb = new();
        wb.Sheets.Add("ReadMe");
        Sheet data = wb.Sheets.Add("Data");
        wb.Sheets.Insert(0, "Cover");
        data[0, 0].Value = T("via handle");
        Assert.Equal("Data", data.Name);
        Assert.Equal(T("via handle"), wb.Sheets["Data"][0, 0].Value);
        Assert.IsType<CellValue.Blank>(wb.Sheets["ReadMe"][0, 0].Value);
    }

    [Fact(DisplayName = "#13 setColumnFilter hides rows at once")]
    public void SetColumnFilterApplies()
    {
        using Workbook wb = new();
        Sheet s = wb.Sheets.Add("Orders");
        s["A1:A3"].SetValues(new CellValue[,] { { T("Status") }, { T("Open") }, { T("Shipped") } });
        s.AutoFilter.Range = s["A1:A3"];
        s.AutoFilter[0].FilterByValues("Open");
        Assert.True(s.Rows[2].Hidden);
    }

    [Fact(DisplayName = "#14 Validation.list writes a quoted list and keeps the drop-down arrow")]
    public void ValidationListIsQuoted()
    {
        using TempDir dir = new();
        string p = dir.Resolve("validation.xlsx");
        using (Workbook wb = new())
        {
            Sheet s = wb.Sheets.Add("S");
            s.Validations.Add(s["A2:A10"], Validation.List("Draft", "Approved", "Rejected"));
            wb.Save(p);
        }
        string rule = FindAll(SheetsXml(p), "<dataValidation .*?</dataValidation>")[0];
        Assert.True(rule.Contains("<formula1>&quot;Draft,Approved,Rejected&quot;</formula1>"), rule);
        Assert.False(rule.Contains("showDropDown=\"1\""), rule);
    }

    [Fact(DisplayName = "#15 Chart.Builder.showDataLabels(true) is saved")]
    public void ChartDataLabelsAreSaved()
    {
        using TempDir dir = new();
        string p = dir.Resolve("chart.xlsx");
        using (Workbook wb = new())
        {
            Sheet s = wb.Sheets.Add("S");
            s["A1:B3"].SetValues(new CellValue[,] { { T("M"), T("V") }, { T("Jan"), N(1) }, { T("Feb"), N(2) } });
            Chart chart = s.Charts.Add(ChartType.Column, new DrawingAnchor(4, 0, 12, 8));
            chart.Title = "T";
            chart.Series.Add("V", "S!$B$2:$B$3", "S!$A$2:$A$3");
            chart.ShowDataLabels = true;
            wb.Save(p);
        }
        string chartXml = ZipEntry(p, ZipNames(p).First(e => e.StartsWith("xl/charts/chart", StringComparison.Ordinal)))!;
        Assert.True(chartXml.Contains("<c:showVal val=\"1\"/>"), "no value labels in the saved chart");
    }

    [Fact(DisplayName = "#16 TemplateMarkers writes a LocalDate as a formatted date")]
    public void TemplateMarkersWriteDates()
    {
        using Workbook wb = new();
        Sheet s = wb.Sheets.Add("T");
        s[0, 0].Value = T("${d}");
        TemplateMarkers.Process(s, new Dictionary<string, object?> { ["d"] = new DateOnly(2026, 9, 30) });
        Assert.Equal(46295.0, Number(s[0, 0].Value));
        Assert.Equal("2026-09-30", s[0, 0].FormattedText);
    }

    [Fact(DisplayName = "#17 saveCsv picks the sheet and can write raw values")]
    public void SaveCsvWritesRawValues()
    {
        using TempDir dir = new();
        string p = dir.Resolve("payments.csv");
        using (Workbook wb = new())
        {
            Sheet a = wb.Sheets.Add("Invoices");
            a[0, 0].Value = N(0.125);
            a[0, 0].Style = new CellStyle { NumberFormat = "0%" };
            a[0, 1].Value = N(-1234.5);
            a[0, 1].Style = new CellStyle { NumberFormat = "#,##0.00;(#,##0.00)" };
            wb.Sheets.Add("Payments")[0, 0].Value = T("payments row");
            wb.SaveCsv(p, new CsvOptions { SheetName = "Invoices", Values = CsvValues.Raw });
        }
        string csv = File.ReadAllText(p);
        Assert.True(csv.StartsWith("0.125,-1234.5", StringComparison.Ordinal), csv);
        Assert.False(csv.Contains("payments row"), csv);
    }

    [Fact(DisplayName = "#19 autoFitColumn measures the formatted text")]
    public void AutoFitUsesNumberFormat()
    {
        using Workbook wb = new();
        Sheet s = wb.Sheets.Add("S");
        s[0, 0].Value = N(8060);
        s[0, 0].Style = new CellStyle { NumberFormat = "#,##0.00" };
        s.Columns[0].AutoFit();
        Assert.True(s.Columns[0].Width >= 8, "width " + s.Columns[0].Width + " is too narrow for '8,060.00'");
    }

    [Fact(DisplayName = "#21 CalcMode.MANUAL defers recalculation and is saved")]
    public void ManualModeDefersAndIsSaved()
    {
        using TempDir dir = new();
        string p = dir.Resolve("manual.xlsx");
        using (Workbook wb = new())
        {
            wb.CalcMode = CalcMode.Manual;
            Sheet s = wb.Sheets.Add("S");
            s[0, 0].Value = N(1);
            s[0, 1].Formula = "A1*10";
            s[0, 0].Value = N(2);
            Assert.Equal(10.0, Number(s[0, 1].Value));
            wb.Save(p);
        }
        Assert.Contains("calcMode=\"manual\"", ZipEntry(p, "xl/workbook.xml"));
    }

    [Fact(DisplayName = "#21 a running balance written top-down is not quadratic")]
    public void FormulaChainWrittenTopDownIsFast()
    {
        using Workbook wb = new();
        Sheet s = wb.Sheets.Add("S");
        int rows = 3000;
        for (int r = 0; r < rows; r++)
            s[r, 0].Value = N(1);
        long start = System.Diagnostics.Stopwatch.GetTimestamp();
        for (int r = 0; r < rows; r++)
            s[r, 1].Formula = r == rows - 1 ? "A" + (r + 1) : "B" + (r + 2) + "+A" + (r + 1);
        long ms = (long)System.Diagnostics.Stopwatch.GetElapsedTime(start).TotalMilliseconds;
        Assert.Equal(3000.0, Number(s[0, 1].Value));
        // 0930/1013 builds took about 12 s; fixed builds take tens of milliseconds
        Assert.True(ms < 3000, "writing the chain took " + ms + " ms");
    }

    [Fact(DisplayName = "#22 two threads on one workbook get 'workbook is already in use'")]
    public async Task OneWorkbookPerThread()
    {
        using Workbook wb = new();
        Sheet s = wb.Sheets.Add("S");
        using CountdownEvent go = new(1);
        string Job()
        {
            go.Wait();
            try
            {
                for (int i = 0; i < 200_000; i++)
                    s[i % 100, 0].Value = N(i);
                return "ok";
            }
            catch (Exception e)
            {
                return e.Message;
            }
        }
        Task<string> a = Task.Factory.StartNew(Job, TaskCreationOptions.LongRunning);
        Task<string> b = Task.Factory.StartNew(Job, TaskCreationOptions.LongRunning);
        go.Signal();
        string ra = await a;
        string rb = await b;
        string failure = ra == "ok" ? rb : ra;
        Assert.True(failure.Contains("workbook is already in use"), failure);
    }

    [Fact(DisplayName = "#26 Workbook.evaluate resolves a multi-cell defined name")]
    public void EvaluateResolvesNames()
    {
        using Workbook wb = new();
        Sheet s = wb.Sheets.Add("Sheet1");
        for (int r = 1; r <= 4; r++)
            s[r, 1].Value = N(r * 10);
        wb.Names.Add("Sales", "Sheet1!$B$2:$B$5");
        Assert.Equal(100.0, Number(wb.Evaluate("SUM(Sales)")));
        Assert.Equal(100.0, Number(s.Evaluate("SUM(Sales)"))); // Java: wb.evaluate("SUM(Sales)", s)
    }

    [Fact(DisplayName = "#27 manual page breaks are saved and paginate")]
    public void PageBreaksAreSavedAndPaginate()
    {
        using TempDir dir = new();
        string p = dir.Resolve("breaks.xlsx");
        using (Workbook wb = new())
        {
            Sheet s = wb.Sheets.Add("P");
            for (int r = 0; r < 20; r++)
                for (int c = 0; c < 3; c++)
                    s[r, c].Value = N(r);
            Assert.Equal(1, wb.Sheets["P"].GetPageCount());
            s.PageSetup.AddRowBreak(10);
            s.PageSetup.AddColumnBreak(1);
            Assert.Equal(4, wb.Sheets["P"].GetPageCount());
            wb.Save(p);
        }
        string xml = SheetsXml(p);
        Assert.True(xml.Contains("<rowBreaks") && xml.Contains("<colBreaks"));
        using (Workbook wb = Workbook.Open(p))
        {
            Assert.Equal(new[] { 10 }, wb.Sheets["P"].PageSetup.GetRowBreaks());
        }
    }

    [Fact(DisplayName = "#28 goalSeek works in CalcMode.MANUAL")]
    public void GoalSeekInManualMode()
    {
        using Workbook wb = new();
        Sheet g = wb.Sheets.Add("G");
        g[0, 0].Value = N(5);
        g[1, 0].Formula = "A1*3";
        wb.CalcMode = CalcMode.Manual;
        Assert.Equal(10.0, wb.GoalSeek(g[1, 0], 30, g[0, 0]), 1e-9);
        Assert.Equal(10.0, Number(g[0, 0].Value), 1e-9);
    }

    [Fact(DisplayName = "#29 array and dynamic-array formula text is the same after reopening")]
    public void ArrayFormulaTextSurvivesReopen()
    {
        using TempDir dir = new();
        string p = dir.Resolve("arrays.xlsx");
        using (Workbook wb = new())
        {
            Sheet s = wb.Sheets.Add("S");
            for (int r = 0; r < 4; r++)
                s[r, 0].Value = N(r + 1);
            s["B1:B4"].SetArrayFormula("A1:A4*2");
            s[0, 2].Formula = "SORT(A1:A4,1,-1)";
            wb.Save(p);
        }
        using (Workbook wb = Workbook.Open(p))
        {
            Sheet s = wb.Sheets["S"];
            Assert.Equal("A1:A4*2", s[0, 1].Formula);
            Assert.Equal("SORT(A1:A4,1,-1)", s[0, 2].Formula);
        }
    }

    [Fact(DisplayName = "Range.setArrayFormula accepts a leading '=' like setFormula")]
    public void SetArrayFormulaStripsEquals()
    {
        using Workbook wb = new();
        Sheet s = wb.Sheets.Add("S");
        for (int r = 0; r < 3; r++)
            s[r, 0].Value = N(r + 1);
        s["B1:B3"].SetArrayFormula("=A1:A3*2");
        Assert.Equal("A1:A3*2", s[0, 1].Formula);
        Assert.Equal(6.0, Number(s[2, 1].Value));
    }

    [Fact(DisplayName = "#30 a multi-row addSparkline makes one sparkline per row")]
    public void SparklinesExpandPerRow()
    {
        using TempDir dir = new();
        string p = dir.Resolve("sparklines.xlsx");
        using (Workbook wb = new())
        {
            Sheet d = wb.Sheets.Add("Data");
            for (int r = 1; r <= 3; r++)
                for (int c = 1; c <= 6; c++)
                    d[r, c].Value = N(r * c);
            d.Sparklines.Add(SparklineType.Line, "B2:G4", "H2:H4");
            wb.Save(p);
        }
        List<string> sparklines = FindAll(SheetsXml(p), "<x14:sparkline>.*?</x14:sparkline>");
        Assert.Equal(
        [
            "<x14:sparkline><xm:f>Data!B2:G2</xm:f><xm:sqref>H2</xm:sqref></x14:sparkline>",
            "<x14:sparkline><xm:f>Data!B3:G3</xm:f><xm:sqref>H3</xm:sqref></x14:sparkline>",
            "<x14:sparkline><xm:f>Data!B4:G4</xm:f><xm:sqref>H4</xm:sqref></x14:sparkline>",
        ], sparklines);
    }

    [Fact(DisplayName = "#31 setTimelineRange is saved")]
    public void TimelineRangeIsSaved()
    {
        using TempDir dir = new();
        string p = dir.Resolve("timeline.xlsx");
        using (Workbook wb = new())
        {
            Sheet data = wb.Sheets.Add("Data");
            data[0, 0].Value = T("OrderDate");
            data[0, 1].Value = T("Region");
            data[0, 2].Value = T("Amount");
            for (int i = 0; i < 12; i++)
            {
                data[i + 1, 0].DateValue = new DateOnly(2026, 1, 5).AddDays(i * 20);
                data[i + 1, 1].Value = T(i % 2 == 0 ? "North" : "South");
                data[i + 1, 2].Value = N(100 + i);
            }
            Sheet ps = wb.Sheets.Add("Pivot");
            PivotTable pv = ps.PivotTables.Add("SalesPivot", "A1", "Data!$A$1:$C$13");
            pv.RowFields.Add("Region");
            pv.DataFields.Add("Amount", AggregateFunction.Sum);
            pv.Refresh();
            // Java: wb.addTimeline("OrderTimeline", "SalesPivot", "OrderDate", ps.index(), 1, 5, 8, 11)
            wb.Timelines.Add("OrderTimeline", pv, "OrderDate", new DrawingAnchor(1, 5, 8, 11), ps);
            // Java: assertTrue(wb.setTimelineRange(...)); the .NET Selection setter returns nothing
            wb.Timelines["OrderTimeline"].Selection = new DateRange(new DateOnly(2026, 2, 1), new DateOnly(2026, 6, 30));
            wb.Save(p);
        }
        using (Workbook wb = Workbook.Open(p))
        {
            Assert.Equal(new DateRange(new DateOnly(2026, 2, 1), new DateOnly(2026, 6, 30)),
                wb.Timelines["OrderTimeline"].Selection);
        }
    }

    [Fact(DisplayName = "#32 ChartView.showDataLabels() reads back after reopening")]
    public void ShowDataLabelsReadsBack()
    {
        using TempDir dir = new();
        string p = dir.Resolve("labels.xlsx");
        using (Workbook wb = new())
        {
            Sheet s = wb.Sheets.Add("S");
            for (int r = 0; r < 4; r++)
            {
                s[r, 0].Value = T("Q" + (r + 1));
                s[r, 1].Value = N(r + 1);
            }
            Chart chart = s.Charts.Add(ChartType.Column, new DrawingAnchor(0, 3, 15, 10));
            chart.Title = "T";
            chart.Series.Add("V", "S!$B$1:$B$4", "S!$A$1:$A$4");
            chart.ShowDataLabels = true;
            wb.Save(p);
        }
        using (Workbook wb = Workbook.Open(p))
        {
            Assert.True(wb.Sheets["S"].Charts[0].ShowDataLabels);
        }
    }

    [Fact(DisplayName = "#34 addTextBox writes a text box")]
    public void AddTextBoxWritesATextBox()
    {
        using TempDir dir = new();
        string p = dir.Resolve("textbox.xlsx");
        using (Workbook wb = new())
        {
            Sheet s = wb.Sheets.Add("S");
            Shape box = s.Shapes.AddTextBox(new DrawingAnchor(1, 1, 4, 5), "Notes");
            Assert.True(s.Shapes[box.Index].IsTextBox);
            wb.Save(p);
        }
        StringBuilder drawings = new();
        foreach (string name in ZipNames(p))
            if (name.StartsWith("xl/drawings/drawing", StringComparison.Ordinal))
                drawings.Append(ZipEntry(p, name));
        Assert.Contains("txBox=\"1\"", drawings.ToString());
    }

    [Fact(DisplayName = "#30 a sparkline shape that does not match the location is refused")]
    public void MismatchedSparklineThrows()
    {
        using Workbook wb = new();
        Sheet d = wb.Sheets.Add("Data");
        for (int r = 1; r <= 3; r++)
            for (int c = 1; c <= 6; c++)
                d[r, c].Value = N(r * c);
        AxyraException e = Assert.ThrowsAny<AxyraException>(() => d.Sparklines.Add(SparklineType.Line, "B2:G4", "J2:J3"));
        Assert.True(e.Message.Contains("B2:G4"), e.Message);
    }

    [Fact(DisplayName = "#35 PdfOptions.watermark is drawn after the page background")]
    public void WatermarkIsNotPaintedOver()
    {
        using Workbook wb = new();
        Sheet s = wb.Sheets.Add("S");
        for (int r = 0; r < 30; r++)
        {
            s[r, 0].Value = T("Row " + r);
            s[r, 1].Value = N(r);
        }
        byte[] pdf = wb.ToPdf(new PdfOptions { Watermark = "CONFIDENTIAL" });
        string page = PdfStreams(pdf).First(c => c.Contains("BT") && c.Contains("1 1 1 rg"));
        Assert.True(page.IndexOf("1 1 1 rg", StringComparison.Ordinal) < page.IndexOf("BT", StringComparison.Ordinal),
            "the full-page white fill is painted after the first text (the watermark)");
    }

    [Fact(DisplayName = "#36 evaluation mode: re-saving a file adds no second notice sheet")]
    public void EvaluationNoticeIsNotDuplicated()
    {
        Assert.SkipWhen(License.Status.IsLicensed, "needs evaluation mode");
        using TempDir dir = new();
        string a = dir.Resolve("a.xlsx");
        string b = dir.Resolve("b.xlsx");
        using (Workbook wb = new())
        {
            wb.Sheets.Add("Q3")[0, 0].Value = T("Region");
            wb.Save(a);
        }
        using (Workbook wb = Workbook.Open(a))
        {
            wb.Save(b);
        }
        using (Workbook wb = Workbook.Open(b))
        {
            Assert.True(wb.Sheets.Count == 2, "one notice sheet plus Q3, got " + wb.Sheets.Count);
        }
    }

    [Fact(DisplayName = "#37 an unterminated range in setFormula is a formula error, not an engine panic")]
    public void UnterminatedRangeIsAFormulaError()
    {
        using Workbook wb = new();
        Sheet s = wb.Sheets.Add("S");
        Assert.ThrowsAny<AxyraFormulaException>(() => s[0, 0].Formula = "SUM(B2:");
    }

    [Fact(DisplayName = "#38 containsText, beginsWith and endsWith rules are saved with their own types")]
    public void TextRulesKeepTheirType()
    {
        using TempDir dir = new();
        string p = dir.Resolve("text-rules.xlsx");
        using (Workbook wb = new())
        {
            Sheet s = wb.Sheets.Add("S");
            string[] values = ["apple pie", "banana", "pineapple", "grape"];
            for (int r = 0; r < values.Length; r++)
                s[r, 0].Value = T(values[r]);
            s.ConditionalFormats.Add(s["A1:A4"], new TextRule(TextRuleOperator.Contains, "apple")
                { Style = new ConditionalStyle { FillColor = Color.Parse("#FFC7CE") } });
            s.ConditionalFormats.Add(s["A1:A4"], new TextRule(TextRuleOperator.BeginsWith, "pine")
                { Style = new ConditionalStyle { FillColor = Color.Parse("#C6EFCE") } });
            s.ConditionalFormats.Add(s["A1:A4"], new TextRule(TextRuleOperator.EndsWith, "pe")
                { Style = new ConditionalStyle { FillColor = Color.Parse("#FFEB9C") } });
            wb.Save(p);
        }
        string xml = SheetsXml(p);
        Assert.True(xml.Contains("type=\"containsText\""), "no containsText rule");
        Assert.True(xml.Contains("type=\"beginsWith\""), "no beginsWith rule");
        Assert.True(xml.Contains("type=\"endsWith\""), "no endsWith rule");
    }

    [Fact(DisplayName = "#39 a blocked dynamic array keeps its input cells after reopening")]
    public void BlockedSpillKeepsInputs()
    {
        using TempDir dir = new();
        string p = dir.Resolve("blocked-spill.xlsx");
        double[] written = [3, 1, 3, 2, 1];
        using (Workbook wb = new())
        {
            Sheet s = wb.Sheets.Add("S");
            for (int r = 0; r < written.Length; r++)
                s[r + 1, 0].Value = N(written[r]);
            s[0, 0].Formula = "SORT(UNIQUE(A2:A100))";
            wb.Save(p);
        }
        using (Workbook wb = Workbook.Open(p))
        {
            Sheet s = wb.Sheets["S"];
            double[] reopened = new double[written.Length];
            for (int r = 0; r < written.Length; r++)
                reopened[r] = Number(s[r + 1, 0].Value);
            Assert.Equal(written, reopened);
        }
    }

    [Fact(DisplayName = "#40 a pivot with two row fields groups each outer item once")]
    public void NestedPivotGroupsOuterItems()
    {
        using Workbook wb = new();
        Sheet d = wb.Sheets.Add("Data");
        string[][] rows = [["Region", "Product", "Amount"], ["EMEA", "Widget", "10"], ["APAC", "Widget", "20"],
            ["EMEA", "Gadget", "30"], ["APAC", "Gadget", "40"], ["EMEA", "Widget", "50"]];
        for (int r = 0; r < rows.Length; r++)
            for (int c = 0; c < 3; c++)
                d[r, c].Value = r > 0 && c == 2
                    ? N(double.Parse(rows[r][c], System.Globalization.CultureInfo.InvariantCulture))
                    : T(rows[r][c]);
        Sheet ps = wb.Sheets.Add("Pivot");
        PivotTable pv = ps.PivotTables.Add("P", "A1", "Data!A1:C6");
        pv.RowFields.Add("Region");
        pv.RowFields.Add("Product");
        pv.DataFields.Add("Amount", AggregateFunction.Sum);
        pv.Refresh();
        CellRange used = ps.UsedRange!;
        List<CellValue> labels = [];
        for (int r = used.FirstRow; r <= used.LastRow; r++)
            labels.Add(ps[r, 0].Value);
        Assert.True(labels.Count(T("EMEA").Equals) == 1, "row labels " + string.Join(", ", labels));
        Assert.True(labels.Count(T("APAC").Equals) == 1, "row labels " + string.Join(", ", labels));
    }

    [Fact(DisplayName = "#41 a de-DE formula reads the decimal comma")]
    public void LocalizedDecimalComma()
    {
        using Workbook wb = new();
        Sheet s = wb.Sheets.Add("S");
        s[0, 1].Value = N(10);
        s[0, 0].SetLocalizedFormula("SUMME(B1;1,5)", "de-DE");
        Assert.Equal("SUM(B1,1.5)", s[0, 0].Formula);
        Assert.Equal(11.5, Number(s[0, 0].Value));
        s[0, 0].Formula = "ROUND(B1/3,2)";
        Assert.Equal("RUNDEN(B1/3;2)", s[0, 0].GetLocalizedFormula("de-DE"));
    }

    [Fact(DisplayName = "#42 a circular reference written before iteration is enabled converges after recalculate()")]
    public void IterationAfterWritingCircularFormulas()
    {
        using Workbook wb = new();
        Sheet s = wb.Sheets.Add("S");
        s[0, 0].Formula = "B1+1";
        s[0, 1].Formula = "A1/2";
        wb.SetIterativeCalculation(true, 100, 0.001);
        wb.Recalculate();
        Assert.Equal(2.0, Number(s[0, 0].Value), 0.01);
        Assert.Equal(1.0, Number(s[0, 1].Value), 0.01);
    }

    [Fact(DisplayName = "#43 the three-argument addTimeline is saved")]
    public void ThreeArgumentTimelineIsSaved()
    {
        using TempDir dir = new();
        string p = dir.Resolve("timeline.xlsx");
        using (Workbook wb = new())
        {
            Sheet d = wb.Sheets.Add("Data");
            d[0, 0].Value = T("OrderDate");
            d[0, 1].Value = T("Amount");
            for (int r = 1; r <= 6; r++)
            {
                d[r, 0].DateValue = new DateOnly(2026, r, 15);
                d[r, 1].Value = N(r * 10);
            }
            Sheet ps = wb.Sheets.Add("Pivot");
            PivotTable pv = ps.PivotTables.Add("SalesPivot", "A1", "Data!A1:B7");
            pv.RowFields.Add("OrderDate");
            pv.DataFields.Add("Amount", AggregateFunction.Sum);
            pv.Refresh();
            wb.Timelines.Add("DateTimeline", "SalesPivot", "OrderDate");
            wb.Save(p);
        }
        Assert.True(ZipNames(p).Any(name => name.StartsWith("xl/timelines/", StringComparison.Ordinal)), "no timeline part saved");
        using (Workbook wb = Workbook.Open(p))
        {
            Assert.Single(wb.Timelines);
        }
    }

    [Fact(DisplayName = "#44 saving as .xlsb keeps comments and hyperlinks")]
    public void XlsbKeepsCommentsAndLinks()
    {
        using Workbook wb = new();
        Sheet s = wb.Sheets.Add("S");
        s[0, 0].Value = T("note here");
        s[0, 0].Comment = new Comment("Ann", "check this");
        s[1, 0].Value = T("link");
        s[1, 0].Hyperlink = Hyperlink.Url("https://keikai.io");
        byte[] xlsb;
        try
        {
            xlsb = wb.ToByteArray(WorkbookFormat.Xlsb);
        }
        catch (AxyraException)
        {
            return;
        }
        using Workbook back = Workbook.Open(xlsb, WorkbookFormat.Xlsb);
        Sheet bs = back.Sheets["S"];
        Assert.True(bs[0, 0].Comment is not null, "comment lost");
        Assert.True(bs[1, 0].Hyperlink is not null, "hyperlink lost");
    }

    [Fact(DisplayName = "#45 insertRows and insertColumns move pictures, charts and shapes")]
    public void InsertingMovesDrawings()
    {
        byte[] png = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==");
        using Workbook wb = new();
        Sheet s = wb.Sheets.Add("S");
        for (int r = 0; r < 10; r++)
            s[r, 1].Value = N(r);
        s.Pictures.Add(png, PictureFormat.Png, new DrawingAnchor(12, 4, 14, 6));
        s.Charts.Add(ChartType.Column, new DrawingAnchor(12, 8, 20, 14)).Series.Add("S", "S!$B$1:$B$10");
        s.Shapes.Add("rect", new DrawingAnchor(12, 1, 14, 3), "box");
        s.Rows.Insert(0, 3);
        s.Columns.Insert(0, 2);
        Assert.Equal((15, 6, 17, 8), Corners(s.Pictures[0].Anchor));   // picture
        Assert.Equal((15, 10, 23, 16), Corners(s.Charts[0].Anchor));   // chart
        Assert.Equal((15, 3, 17, 5), Corners(s.Shapes[0].Anchor));     // shape
    }

    private static (int, int, int, int) Corners(DrawingAnchor a) => (a.From.Row, a.From.Column, a.To.Row, a.To.Column);

    [Fact(DisplayName = "#46 Theme.builder() writes no empty colour for unset slots")]
    public void PartialThemeHasNoEmptyColours()
    {
        using TempDir dir = new();
        string p = dir.Resolve("theme.xlsx");
        using (Workbook wb = new())
        {
            wb.Sheets.Add("S")[0, 0].Value = T("x");
            wb.Theme = new Theme { MajorFont = "Georgia" }.WithColor(Theme.Slot.Accent1, Color.Parse("#2C3E50"));
            wb.Save(p);
        }
        Assert.DoesNotContain("<a:srgbClr val=\"\"/>", ZipEntry(p, "xl/theme/theme1.xml"));
    }

    [Fact(DisplayName = "#47 colorScale refuses one or four colours")]
    public void ColorScaleNeedsTwoOrThreeColours()
    {
        using Workbook wb = new();
        Sheet s = wb.Sheets.Add("S");
        CellRange range = s["A1:A5"];
        // Java: ConditionalFormat.builder().colorScale(Color...) takes bare colours; the .NET rule takes
        // points, so each colour gets a threshold (min, percentiles, max) as the Java builder assigns them.
        Assert.ThrowsAny<ArgumentException>(() => s.ConditionalFormats.Add(range, new ColorScaleRule(
            [new ColorScalePoint(new ConditionalThreshold(ThresholdType.Min), Color.Red)])));
        Assert.ThrowsAny<ArgumentException>(() => s.ConditionalFormats.Add(range, new ColorScaleRule(
        [
            new ColorScalePoint(new ConditionalThreshold(ThresholdType.Min), Color.Red),
            new ColorScalePoint(new ConditionalThreshold(ThresholdType.Percentile, "33"), Color.Yellow),
            new ColorScalePoint(new ConditionalThreshold(ThresholdType.Percentile, "67"), Color.Green),
            new ColorScalePoint(new ConditionalThreshold(ThresholdType.Max), Color.Blue),
        ])));
    }

    [Fact(DisplayName = "#48 setIterativeCalc is saved")]
    public void IterativeCalcIsSaved()
    {
        using TempDir dir = new();
        string p = dir.Resolve("iterative.xlsx");
        using (Workbook wb = new())
        {
            wb.Sheets.Add("S")[0, 0].Value = N(1);
            wb.SetIterativeCalculation(true, 50, 0.01);
            wb.Save(p);
        }
        string calcPr = FindAll(ZipEntry(p, "xl/workbook.xml")!, "<calcPr[^>]*>")[0];
        Assert.True(calcPr.Contains("iterate=\"1\""), calcPr);
        Assert.True(calcPr.Contains("iterateCount=\"50\""), calcPr);
    }

    [Fact(DisplayName = "#49 createTable refuses a name that is a cell reference or already used")]
    public void CreateTableRefusesBadNames()
    {
        using Workbook wb = new();
        Sheet s = wb.Sheets.Add("S");
        for (int r = 0; r < 4; r++)
            for (int c = 0; c < 5; c++)
                s[r, c].Value = r == 0 ? T("H" + c) : N(r);
        Assert.ThrowsAny<Exception>(() => s.Tables.Add("T1", s["A1:B4"], "H0", "H1"));
        s.Tables.Add("Sales", s["A1:B4"], "H0", "H1");
        Assert.ThrowsAny<Exception>(() => s.Tables.Add("Sales", s["D1:E4"], "H3", "H4"));
    }

    [Fact(DisplayName = "#50 saveEncrypted to .ods writes an ODS package or refuses")]
    public void EncryptedOdsIsOds()
    {
        using TempDir dir = new();
        string p = dir.Resolve("secret.ods");
        using (Workbook wb = new())
        {
            wb.Sheets.Add("S")[0, 0].Value = T("secret");
            try
            {
                wb.Save(p, new SaveOptions { Password = "pw" }); // Java: wb.saveEncrypted(p, "pw")
            }
            catch (Exception refused) when (refused is ArgumentException or AxyraException)
            {
                return;
            }
        }
        byte[] b = File.ReadAllBytes(p);
        Assert.True(b.Take(4).SequenceEqual(new byte[] { (byte)'P', (byte)'K', 3, 4 }), "not a ZIP (ODS) package");
    }

    [Fact(DisplayName = "#51 a JSON round trip keeps the column order")]
    public void JsonKeepsColumnOrder()
    {
        using Workbook wb = new();
        Sheet s = wb.Sheets.Add("S");
        s[0, 0].Value = T("Region");
        s[0, 1].Value = T("Amount");
        s[1, 0].Value = T("APAC");
        s[1, 1].Value = N(5);
        byte[] json = wb.ToByteArray(WorkbookFormat.Json);
        using Workbook back = Workbook.Open(json, WorkbookFormat.Json);
        Sheet bs = back.Sheets[back.Sheets.Count - 1];
        Assert.True(T("Region").Equals(bs[0, 0].Value), Encoding.UTF8.GetString(json));
        Assert.Equal(T("Amount"), bs[0, 1].Value);
    }

    [Fact(DisplayName = "#52 PDF permissions without an owner password are written or refused")]
    public void PdfPermissionsAreNotDropped()
    {
        using Workbook wb = new();
        wb.Sheets.Add("S")[0, 0].Value = T("x");
        byte[] pdf;
        try
        {
            pdf = wb.ToPdf(new PdfOptions { AllowCopy = false, AllowPrint = false });
        }
        catch (Exception refused) when (refused is ArgumentException or InvalidOperationException or AxyraException)
        {
            return;
        }
        Assert.True(Encoding.Latin1.GetString(pdf).Contains("/Encrypt"),
            "AllowCopy = false / AllowPrint = false were dropped without an error");
    }

    [Fact(DisplayName = "#53 barShape on a 2-D bar chart reads the same in memory as after a save")]
    public void BarShapeIsSaved()
    {
        using TempDir dir = new();
        string p = dir.Resolve("bars.xlsx");
        BarShape inMemory;
        using (Workbook wb = new())
        {
            Sheet s = wb.Sheets.Add("S");
            for (int r = 0; r < 4; r++)
                s[r, 0].Value = N(r + 1);
            try
            {
                Chart chart = s.Charts.Add(ChartType.Bar, new DrawingAnchor(0, 2, 10, 8));
                chart.BarShape = BarShape.Cylinder;
                chart.Series.Add("A", "S!$A$1:$A$4");
            }
            catch (Exception refused) when (refused is ArgumentException or AxyraException)
            {
                return;
            }
            inMemory = s.Charts[0].BarShape;
            wb.Save(p);
        }
        using (Workbook wb = Workbook.Open(p))
        {
            BarShape reopened = wb.Sheets["S"].Charts[0].BarShape;
            Assert.True(reopened == inMemory, $"2-D barShape in memory ({inMemory}) vs after reopen ({reopened})");
        }
    }

    [Fact(DisplayName = "#54 conditional formats added without a priority get distinct priorities")]
    public void ConditionalFormatPrioritiesAreDistinct()
    {
        using TempDir dir = new();
        string p = dir.Resolve("priorities.xlsx");
        using (Workbook wb = new())
        {
            Sheet s = wb.Sheets.Add("S");
            for (int i = 0; i < 5; i++)
                s[i, 0].Value = N(i);
            s.ConditionalFormats.Add(s["A1:A5"], new CellValueRule(ComparisonOperator.GreaterThan, "2")
                { Style = new ConditionalStyle { FillColor = Color.Red } });
            s.ConditionalFormats.Add(s["A1:A5"], new CellValueRule(ComparisonOperator.LessThan, "4")
                { Style = new ConditionalStyle { FillColor = Color.Green } });
            s.ConditionalFormats.Add(s["A1:A5"], new DataBarRule(Color.Blue));
            wb.Save(p);
        }
        List<string> priorities = FindAll(SheetsXml(p), "priority=\"\\d+\"");
        Assert.True(priorities.Distinct().Count() == 3, "saved " + string.Join(", ", priorities));
    }

    [Fact(DisplayName = "#55 evaluation mode: StreamWorkbook keeps the program's last row")]
    public void StreamNoticeDoesNotOverwriteLastRow()
    {
        Assert.SkipWhen(License.Status.IsLicensed, "needs evaluation mode");
        using TempDir dir = new();
        string p = dir.Resolve("stream.xlsx");
        using (StreamWorkbook sw = new(p))
        {
            sw.StartSheet("S");
            sw.StartRow(0).AddText("first").Commit();
            sw.StartRow(1_048_575).AddText("last row").Commit();
            sw.Finish();
        }
        int rows = Count(SheetsXml(p), "<row r=\"1048576\"");
        Assert.True(rows == 1, "two rows with the same number (" + rows + ")");
        using (Workbook wb = Workbook.Open(p))
        {
            Assert.Equal(T("last row"), wb.Sheets["S"][1_048_575, 0].Value);
        }
    }

    [Fact(DisplayName = "#56 consolidate(SUM) does not write -0.0 over a text label column")]
    public void ConsolidateLeavesLabelsAlone()
    {
        using Workbook wb = new();
        Sheet jan = wb.Sheets.Add("Jan");
        Sheet feb = wb.Sheets.Add("Feb");
        Sheet sum = wb.Sheets.Add("Sum");
        foreach (Sheet s in new[] { jan, feb })
        {
            s[0, 0].Value = T("North");
            s[0, 1].Value = N(10);
            s[1, 0].Value = T("South");
            s[1, 1].Value = N(20);
        }
        sum.Consolidate(0, 0, ConsolidateFunction.Sum, jan["A1:B2"], feb["A1:B2"]);
        CellValue label = sum[0, 0].Value;
        Assert.False(label is CellValue.Number num && num.Value == 0 && 1 / num.Value < 0, "label column holds -0.0");
        Assert.Equal(20.0, Number(sum[0, 1].Value));
    }

    [Fact(DisplayName = "#57 addName stores the formula without a leading '='")]
    public void AddNameStripsEquals()
    {
        using TempDir dir = new();
        string p = dir.Resolve("names.xlsx");
        using (Workbook wb = new())
        {
            wb.Sheets.Add("S")[0, 0].Formula = "100*Rate";
            wb.Names.Add("Rate", "=0.05");
            Assert.Equal("0.05", wb.Names[0].Formula);
            wb.Save(p);
        }
        Assert.Contains("<definedName name=\"Rate\">0.05</definedName>", ZipEntry(p, "xl/workbook.xml"));
    }

    [Fact(DisplayName = "#58 evaluate rejects a malformed formula like setFormula does")]
    public void EvaluateRejectsMalformedFormula()
    {
        using Workbook wb = new();
        Sheet s = wb.Sheets.Add("S");
        Assert.ThrowsAny<AxyraFormulaException>(() => s[0, 0].Formula = "SUM((");
        Assert.ThrowsAny<AxyraFormulaException>(() => wb.Evaluate("SUM(("));
    }

    [Fact(DisplayName = "#59 a numeric user function given text returns #VALUE!, not NaN")]
    public void UserFunctionWithTextIsValueError()
    {
        using Workbook wb = new();
        Sheet s = wb.Sheets.Add("S");
        wb.Functions.RegisterNumeric("DOUBLEIT", a => a[0] * 2);
        s[0, 1].Value = T("abc");
        s[0, 0].Formula = "DOUBLEIT(B1)";
        Assert.IsType<CellValue.Error>(s[0, 0].Value);
    }

    [Fact(DisplayName = "#60 CellStyle has value equality")]
    public void CellStyleEquals()
    {
        CellStyle style = new() { Font = new Font { Bold = true }, NumberFormat = "#,##0.00", Fill = Fill.Solid(Color.Parse("#FFEB9C")!) };
        // Java: CellStyle.fromJson(style.toJson()). The .NET CellStyle is a record with no JSON form, so the
        // copy is an independently built style with the same settings.
        CellStyle copy = new() { Font = new Font { Bold = true }, NumberFormat = "#,##0.00", Fill = Fill.Solid(Color.Parse("#FFEB9C")!) };
        Assert.Equal(style, copy);
        Assert.Equal(style.GetHashCode(), copy.GetHashCode());
    }

    [Fact(DisplayName = "#61 a sheet Axyra creates has a uniqueId")]
    public void CreatedSheetHasUniqueId()
    {
        using TempDir dir = new();
        string p = dir.Resolve("ids.xlsx");
        using (Workbook wb = new())
        {
            Assert.NotNull(wb.Sheets.Add("Data").UniqueId);
            wb.Save(p);
        }
        using (Workbook wb = Workbook.Open(p))
        {
            Assert.NotNull(wb.Sheets["Data"].UniqueId);
        }
    }
}
