using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using Axyra.Sheets.Calc;
using Axyra.Sheets.Drawing;
using Axyra.Sheets.Formula;
using Axyra.Sheets.IO;
using Axyra.Sheets.Pivots;
using Axyra.Sheets.Style;
using static Axyra.Sheets.Examples.Tests.Regression.IssueTestSupport;

namespace Axyra.Sheets.Examples.Tests.Regression;

/// <summary>
/// Features F36-F52 of the axyra-sheets skill test (CSV, imports, streaming, templates, functions,
/// document properties, calculation, formula notation, named styles and themes, data tools, save
/// formats, signatures, slicers, form controls, split and combine). F47 (licence API) is not tested
/// here: it changes the licence state, which the whole process shares.
/// </summary>
public class FeatureDataTest
{
    /// <summary>F36 import CSV text, split a semicolon column, save a sheet as raw CSV</summary>
    [Fact]
    public void CsvImportAndExport()
    {
        using TempDir dir = new();
        byte[] csv = Encoding.UTF8.GetBytes("Product,Qty\nApple,10\n\"Banana, ripe\",25\n");
        using (Workbook wb = Workbook.Open(csv, WorkbookFormat.Csv))
        {
            Sheet s = wb.Sheets[wb.Sheets.Count - 1];
            Assert.Equal(T("Banana, ripe"), s[2, 0].Value);
            Assert.Equal(25.0, Number(s[2, 1].Value));
        }
        string output = dir.Resolve("semicolon.csv");
        using (Workbook wb = new())
        {
            Sheet s = wb.Sheets.Add("Semicolon");
            s[0, 0].Value = T("Name;City;Score");
            s[1, 0].Value = T("Ann;Paris;91");
            Assert.Equal(3, s["A1:A2"].TextToColumns(";"));
            Assert.Equal(91.0, Number(s[1, 2].Value));
            wb.SaveCsv(output, new CsvOptions { SheetName = "Semicolon", Values = CsvValues.Raw });
        }
        Assert.True(File.ReadAllText(output).StartsWith("Name,City,Score", StringComparison.Ordinal), File.ReadAllText(output));
    }

    // Java: record Item(String sku, String name, double price) with columns "sku", "name", "price".
    // .NET records have PascalCase properties, so the column names are the property names.
    public record Item(string Sku, string Name, double Price);

    /// <summary>F37 import .NET records at a cell with a header row (Java: Java records)</summary>
    [Fact]
    public void ImportJavaData()
    {
        using Workbook wb = new();
        Sheet s = wb.Sheets.Add("S");
        List<Item> items = [new Item("00042", "Widget", 2.5), new Item("00043", "Gadget", 12.75)];
        int rows = s.ImportData(1, 1, items, new ImportOptions
        {
            Columns = ["Sku", "Name", "Price"],
            ColumnLabels = ["SKU", "Name", "Price"],
        });
        Assert.Equal(3, rows);
        Assert.Equal(T("SKU"), s[1, 1].Value);
        Assert.Equal(T("00042"), s[2, 1].Value);
        Assert.Equal(12.75, Number(s[3, 3].Value));
    }

    /// <summary>F38 stream 100,000 rows x 5 columns with a bold header, reopen the last row</summary>
    [Fact]
    public void StreamingWrite()
    {
        using TempDir dir = new();
        string p = dir.Resolve("big.xlsx");
        using (StreamWorkbook sw = new(p))
        {
            int bold = sw.AddStyle(StreamStyle.Default with { Bold = true });
            sw.StartSheet("Data");
            StreamRow header = sw.StartRow(0);
            foreach (string h in new[] { "ID", "Name", "Qty", "Price", "Amount" })
                header.AddText(h, bold);
            header.Commit();
            for (int r = 1; r <= 100_000; r++)
                sw.StartRow(r).AddNumber(r).AddText("Item-" + r).AddNumber(1).AddNumber(0.5).AddNumber(0.5).Commit();
            sw.Finish();
        }
        Assert.True(new FileInfo(p).Length > 1_000_000);
        using (Workbook wb = Workbook.Open(p))
        {
            Sheet s = wb.Sheets["Data"];
            Assert.Equal(100_000.0, Number(s[100_000, 0].Value));
            Assert.Equal(T("Item-100000"), s[100_000, 1].Value);
            Assert.True(s[0, 0].Style.Font.Bold);
        }
    }

    /// <summary>F39 template markers expand a list of maps into rows</summary>
    [Fact]
    public void TemplateMarkers()
    {
        using Workbook wb = new();
        Sheet s = wb.Sheets.Add("T");
        s[0, 0].Value = T("${items.name}");
        s[0, 1].Value = T("${items.price}");
        s[1, 0].Value = T("Total");
        s[1, 1].Value = T("${items.price:sum}");
        Template.TemplateMarkers.Process(s, new Dictionary<string, object?>
        {
            ["items"] = new List<Dictionary<string, object>>
            {
                new() { ["name"] = "Apple", ["price"] = 1.25 },
                new() { ["name"] = "Banana", ["price"] = 0.5 },
                new() { ["name"] = "Cherry", ["price"] = 4.0 },
            },
        });
        Assert.Equal(T("Apple"), s[0, 0].Value);
        Assert.Equal(T("Cherry"), s[2, 0].Value);
        Assert.Equal(T("Total"), s[3, 0].Value);
        Assert.Equal(5.75, Number(s[3, 1].Value));
    }

    /// <summary>F40 a numeric and a typed user-defined function</summary>
    [Fact]
    public void CustomFunctions()
    {
        using Workbook wb = new();
        Sheet s = wb.Sheets.Add("S");
        wb.Functions.RegisterNumeric("DOUBLEIT", a => a[0] * 2);
        wb.Functions.Register("GREET", argv => CellValue.Of("Hello, "
            + (argv[0] is CellValue.Text name ? name.Value : "?") + "!"));
        s[0, 0].Value = N(21);
        s[0, 1].Formula = "DOUBLEIT(A1)";
        s[0, 2].Formula = "GREET(\"Alice\")";
        wb.Recalculate();
        Assert.Equal(42.0, Number(s[0, 1].Value));
        Assert.Equal(T("Hello, Alice!"), s[0, 2].Value);
        Assert.Equal(100.0, Number(wb.Evaluate("DOUBLEIT(50)")));
    }

    /// <summary>F41 the function catalog</summary>
    [Fact]
    public void FunctionCatalog()
    {
        Assert.True(BuiltInFunctions.Count > 400, "count " + BuiltInFunctions.Count);
        Assert.True(BuiltInFunctions.Contains("XLOOKUP"));
        Assert.False(BuiltInFunctions.Contains("MADEUPFUNC"));
        Assert.Equal("ABS", BuiltInFunctions.Names[0]);
    }

    /// <summary>F42 document and custom properties survive save and reopen</summary>
    [Fact]
    public void DocumentProperties()
    {
        using TempDir dir = new();
        string p = dir.Resolve("props.xlsx");
        using (Workbook wb = new())
        {
            wb.Sheets.Add("S")[0, 0].Value = T("x");
            wb.Title = "Quarterly Sales Report";
            wb.Author = "Jane Doe";
            wb.Subject = "Q3 figures";
            wb.Keywords = "sales, quarterly";
            wb.SetCustomProperty("Department", "Finance");
            wb.SetCustomProperty("Budget", 125000.5);
            wb.Save(p);
        }
        using (Workbook wb = Workbook.Open(p))
        {
            Assert.Equal("Quarterly Sales Report", wb.Title);
            Assert.Equal("Jane Doe", wb.Author);
            Assert.Equal("Q3 figures", wb.Subject);
            Assert.Equal("sales, quarterly", wb.Keywords);
            List<string> names = wb.CustomProperties.Select(x => x.Name).ToList();
            Assert.True(new[] { "Department", "Budget" }.All(names.Contains), string.Join(", ", names));
        }
    }

    /// <summary>F43 manual calculation, dirty recalculation, precedents, dependents, goal seek</summary>
    [Fact]
    public void CalculationControl()
    {
        using Workbook wb = new();
        Sheet s = wb.Sheets.Add("Calc");
        s[0, 1].Value = N(10);
        s[1, 1].Value = N(5);
        s[2, 1].Formula = "B1*B2";
        s[3, 1].Formula = "B3*1.1";
        wb.CalcMode = CalcMode.Manual;
        s[0, 1].Value = N(20);
        Assert.Equal(50.0, Number(s[2, 1].Value));
        wb.RecalculateDirty();
        Assert.Equal(100.0, Number(s[2, 1].Value));
        // Java: wb.precedents(s, 2, 1) as CellRef.a1(); .NET: cell.Precedents as Cell.Address
        List<string> precedents = s[2, 1].Precedents.Select(c => c.Address).ToList();
        Assert.True(precedents.Any(a => a.EndsWith("B1", StringComparison.Ordinal))
            && precedents.Any(a => a.EndsWith("B2", StringComparison.Ordinal)), string.Join(", ", precedents));
        Assert.Contains(s[2, 1].Dependents.Select(c => c.Address), a => a.EndsWith("B4", StringComparison.Ordinal));
        Assert.Equal(40.0, wb.GoalSeek(s[3, 1], 220, s[0, 1]), 1e-6);
    }

    /// <summary>F44 A1 to R1C1 and back with an anchor at C3; a German localized formula</summary>
    [Fact]
    // known issue: CONVERSION-NOTES/regression-features.md, "FormulaToR1C1 / FormulaToA1 anchor is 1-based"
    [Trait("Category", "known-issue")]
    public void FormulaNotation()
    {
        using Workbook wb = new();
        // Java passes the 1-based anchor (3, 3); the .NET anchor is documented as zero-based like
        // sheet[row, column] (docs/dotnet-api/12, "A1 ⇄ R1C1 conversion"), so C3 is (2, 2).
        string r1c1 = wb.FormulaToR1C1("A1+$B$2+SUM(A1:A3)", 2, 2);
        Assert.Equal("R[-2]C[-2]+R2C2+SUM(R[-2]C[-2]:RC[-2])", r1c1);
        Assert.Equal("A1+$B$2+SUM(A1:A3)", wb.FormulaToA1(r1c1, 2, 2));
        Sheet s = wb.Sheets.Add("S");
        for (int r = 0; r < 3; r++)
            s[r, 0].Value = N((r + 1) * 10);
        s[4, 0].SetLocalizedFormula("SUMME(A1:A3)", "de-DE");
        Assert.Equal("SUM(A1:A3)", s[4, 0].Formula);
        Assert.Equal(60.0, Number(s[4, 0].Value));
    }

    /// <summary>F45 a named style applied and updated, default font and theme, reopened</summary>
    [Fact]
    public void NamedStylesAndThemes()
    {
        using TempDir dir = new();
        string p = dir.Resolve("styles.xlsx");
        using (Workbook wb = new())
        {
            Sheet s = wb.Sheets.Add("S");
            s[0, 0].Value = T("Header");
            wb.RegisterNamedStyle("ReportHeader", new CellStyle { Font = new Font { Bold = true, Size = 12 } });
            s["A1:B1"].ApplyNamedStyle("ReportHeader");
            Assert.True(wb.UpdateNamedStyle("ReportHeader",
                new CellStyle { Font = new Font { Bold = true, Italic = true, Size = 14 } }));
            s["A1:B1"].ApplyNamedStyle("ReportHeader");
            wb.DefaultFont = wb.DefaultStyle.Font with { Name = "Arial", Size = 10 };
            Theme theme = new() { MajorFont = "Georgia", MinorFont = "Arial" };
            string[] hex =
            [
                "000000", "FFFFFF", "1F3864", "E7E6E6", "1F77B4", "FF7F0E", "2CA02C", "D62728",
                "9467BD", "8C564B", "0563C1", "954F72",
            ];
            Theme.Slot[] slots = Enum.GetValues<Theme.Slot>();
            for (int i = 0; i < hex.Length; i++)
                theme = theme.WithColor(slots[i], Color.Parse("#" + hex[i]));
            wb.Theme = theme;
            wb.Save(p);
        }
        using (Workbook wb = Workbook.Open(p))
        {
            Assert.True(wb.HasNamedStyle("ReportHeader"));
            CellStyle header = wb.Sheets["S"][0, 0].Style;
            Assert.True(header.Font.Italic);
            Assert.Equal(14.0, header.Font.Size);
            Assert.Equal("Arial", wb.DefaultStyle.Font.Name);
            Assert.Equal("Georgia", wb.Theme!.MajorFont);
        }
    }

    /// <summary>F46 text to columns, remove duplicates, auto-fill a series, consolidate with SUM</summary>
    [Fact]
    public void DataTools()
    {
        using Workbook wb = new();
        Sheet s = wb.Sheets.Add("S");
        s[0, 0].Value = T("a,b,c");
        s[1, 0].Value = T("d,e,f");
        Assert.Equal(3, s["A1:A2"].TextToColumns(","));
        Assert.Equal(T("f"), s[1, 2].Value);

        Sheet d = wb.Sheets.Add("D");
        string[][] rows = [["Name", "Qty"], ["Apple", "10"], ["Pear", "5"], ["Apple", "10"]];
        for (int r = 0; r < rows.Length; r++)
        {
            d[r, 0].Value = T(rows[r][0]);
            d[r, 1].Value = r == 0 ? T(rows[r][1]) : N(double.Parse(rows[r][1]));
        }
        Assert.Equal(1, d["A1:B4"].RemoveDuplicates([0, 1], true));

        Sheet f = wb.Sheets.Add("F");
        f[0, 0].Value = N(1);
        f[1, 0].Value = N(2);
        f["A1:A2"].AutoFill(f["A1:A10"]);
        Assert.Equal(10.0, Number(f[9, 0].Value));

        Sheet jan = wb.Sheets.Add("Jan");
        Sheet feb = wb.Sheets.Add("Feb");
        Sheet sum = wb.Sheets.Add("Sum");
        foreach (Sheet m in new[] { jan, feb })
            m["A1:B2"].SetValues(new CellValue[,] { { N(1), N(2) }, { N(3), N(4) } });
        sum.Consolidate(0, 0, ConsolidateFunction.Sum, jan["A1:B2"], feb["A1:B2"]);
        Assert.Equal(8.0, Number(sum[1, 1].Value));
    }

    /// <summary>F48 the same workbook saved as xlsx, xlsm, ods, csv, tsv, html and pdf, with formula policies</summary>
    [Fact]
    public void SaveFormats()
    {
        using TempDir dir = new();
        using (Workbook wb = new())
        {
            Sheet s = wb.Sheets.Add("Sales");
            s["A1:A2"].SetValues(new CellValue[,] { { N(1) }, { N(2) } });
            s[2, 0].Formula = "SUM(A1:A2)";
            foreach (string ext in new[] { "xlsx", "xlsm", "ods", "csv", "tsv" })
            {
                string p = dir.Resolve("out." + ext);
                wb.Save(p);
                Assert.True(new FileInfo(p).Length > 0, ext);
            }
            wb.SaveHtml(dir.Resolve("out.html"), new HtmlOptions { SingleFile = true });
            Assert.Contains("<table", File.ReadAllText(dir.Resolve("out.html")));
            wb.SavePdf(dir.Resolve("out.pdf"), new PdfOptions());
            Assert.True(new FileInfo(dir.Resolve("out.pdf")).Length > 0);
            wb.Save(dir.Resolve("values.xlsx"), new SaveOptions { FormulaPolicy = FormulaPolicy.ValuesOnly });
            Assert.ThrowsAny<ArgumentException>(
                () => wb.Save(dir.Resolve("keep.csv"), new SaveOptions { FormulaPolicy = FormulaPolicy.Keep }));
        }
        using (Workbook wb = Workbook.Open(dir.Resolve("values.xlsx")))
        {
            Assert.Null(wb.Sheets["Sales"][2, 0].Formula);
            Assert.Equal(3.0, Number(wb.Sheets["Sales"][2, 0].Value));
        }
    }

    /// <summary>F49 sign a workbook with a key generated at run time (needs a licence)</summary>
    [Fact]
    public void DigitalSignature()
    {
        Assert.SkipUnless(License.Status.IsLicensed, "signing needs a licence; the evaluation build refuses it");
        // Java runs keytool to make a PKCS#12 key pair; .NET makes the same self-signed RSA-2048 /
        // SHA-256 certificate in process with CertificateRequest.
        using RSA rsa = RSA.Create(2048);
        CertificateRequest request = new("CN=Axyra Test Signer", rsa, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
        using X509Certificate2 cert = request.CreateSelfSigned(DateTimeOffset.UtcNow.AddMinutes(-5), DateTimeOffset.UtcNow.AddDays(30));
        byte[] key = rsa.ExportPkcs8PrivateKey();
        byte[] certificate = cert.RawData;
        using TempDir dir = new();
        string signed = dir.Resolve("signed.xlsx");
        using (Workbook wb = new())
        {
            wb.Sheets.Add("Data")[0, 0].Value = N(200);
            wb.Sign(signed, key, certificate);
        }
        using (Workbook wb = Workbook.Open(signed))
        {
            SignatureInfo info = wb.Signatures[0];
            Assert.Equal("Axyra Test Signer", info.SignerName);
            Assert.Equal(SignatureStatus.Valid, info.Status);
        }
    }

    /// <summary>F50 a pivot slicer with a selection and a timeline with a range, reopened</summary>
    [Fact]
    public void SlicersAndTimelines()
    {
        using TempDir dir = new();
        string p = dir.Resolve("slicers.xlsx");
        using (Workbook wb = new())
        {
            Sheet data = wb.Sheets.Add("Data");
            string[] header = ["OrderDate", "Region", "Amount"];
            for (int c = 0; c < 3; c++)
                data[0, c].Value = T(header[c]);
            string[] regions = ["North", "South", "East", "West"];
            for (int i = 0; i < 12; i++)
            {
                data[i + 1, 0].DateValue = new DateOnly(2026, 1, 5).AddDays(i * 20);
                data[i + 1, 1].Value = T(regions[i % 4]);
                data[i + 1, 2].Value = N(100 + i);
            }
            Sheet ps = wb.Sheets.Add("Pivot");
            PivotTable pv = ps.PivotTables.Add("SalesPivot", "A1", "Data!$A$1:$C$13");
            pv.RowFields.Add("Region");
            pv.DataFields.Add("Amount", AggregateFunction.Sum);
            pv.Refresh();
            // Java: addSlicer(name, field, pivot, sheet, 1, 6, 13, 10) / addTimeline(..., 15, 6, 21, 12)
            Slicer slicer = wb.Slicers.Add("RegionSlicer", pv, "Region", new DrawingAnchor(1, 6, 13, 10));
            Timeline timeline = wb.Timelines.Add("OrderTimeline", pv, "OrderDate", new DrawingAnchor(15, 6, 21, 12), ps);
            // Java's setSlicerSelection / setTimelineRange return true; .NET assigns properties
            slicer.SelectedItems = ["North", "East"];
            timeline.Selection = new DateRange(new DateOnly(2026, 2, 1), new DateOnly(2026, 6, 30));
            wb.Save(p);
        }
        using (Workbook wb = Workbook.Open(p))
        {
            Assert.Equal("Region", wb.Slicers[0].SourceField);
            Assert.Equal(["East", "North"], wb.Slicers["RegionSlicer"].SelectedItems.Order().ToList());
            Assert.Equal(new DateRange(new DateOnly(2026, 2, 1), new DateOnly(2026, 6, 30)),
                wb.Timelines["OrderTimeline"].Selection);
        }
    }

    /// <summary>F51 a linked checkbox, a combo box and a button, reopened</summary>
    [Fact]
    public void FormControls()
    {
        using TempDir dir = new();
        string p = dir.Resolve("controls.xlsx");
        using (Workbook wb = new())
        {
            Sheet s = wb.Sheets.Add("S");
            for (int r = 0; r < 3; r++)
                s[r, 4].Value = T("Item " + r);
            FormControl check = s.FormControls.Add(FormControlType.CheckBox, new DrawingAnchor(0, 1, 1, 2), "AgreeCheck");
            check.Text = "I agree";
            check.LinkedCell = "$C$1";
            check.Checked = true;
            FormControl combo = s.FormControls.Add(FormControlType.ComboBox, new DrawingAnchor(2, 1, 3, 3), "FruitCombo");
            combo.ListFillRange = "$E$1:$E$3";
            combo.LinkedCell = "$C$4";
            combo.SelectedIndex = 2;
            FormControl button = s.FormControls.Add(FormControlType.Button, new DrawingAnchor(5, 1, 7, 3), "RunButton");
            button.Text = "Run";
            wb.Save(p);
        }
        using (Workbook wb = Workbook.Open(p))
        {
            FormControlCollection controls = wb.Sheets["S"].FormControls;
            Assert.Equal([FormControlType.CheckBox, FormControlType.ComboBox, FormControlType.Button],
                controls.Select(c => c.Type).ToList());
            Assert.Equal("$C$1", controls[0].LinkedCell);
            Assert.True(controls[0].Checked);
            Assert.Equal("$E$1:$E$3", controls[1].ListFillRange);
        }
    }

    /// <summary>F52 combine two workbooks and split one into a file per sheet</summary>
    [Fact]
    public void SplitAndCombine()
    {
        using TempDir dir = new();
        using Workbook a = new();
        using Workbook b = new();
        a.Sheets.Add("North")[0, 0].Value = N(100);
        a.Sheets.Add("South")[0, 0].Value = N(200);
        b.Sheets.Add("East")[0, 0].Value = N(300);
        a.Combine(b);
        List<string> names = [];
        foreach (Sheet s in a.Sheets)
            names.Add(s.Name);
        Assert.Equal(["North", "South", "East"], names);
        IReadOnlyList<string> files = a.SplitToFiles(dir.Path, "part_{index}_{name}.xlsx");
        Assert.Equal(3, files.Count);
        Assert.Equal("part_2_East.xlsx", Path.GetFileName(files[2]));
        using Workbook east = Workbook.Open(files[2]);
        Assert.Equal(300.0, Number(east.Sheets["East"][0, 0].Value));
    }
}
