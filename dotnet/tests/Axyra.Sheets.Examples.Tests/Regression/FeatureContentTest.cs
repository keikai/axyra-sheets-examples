using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using Axyra.Sheets.Charts;
using Axyra.Sheets.Drawing;
using Axyra.Sheets.IO;
using Axyra.Sheets.Pivots;
using Axyra.Sheets.Style;
using static Axyra.Sheets.Examples.Tests.Regression.IssueTestSupport;

namespace Axyra.Sheets.Examples.Tests.Regression;

/// <summary>
/// Features F24-F35 of the axyra-sheets skill test (charts, sparklines, pictures, shapes, pivots,
/// sheet management, protection, encryption, page setup, PDF and image export).
/// </summary>
public class FeatureContentTest
{
    private const string Notice = "Axyra Evaluation Copy";

    /// <summary>Page objects in a PDF (<c>/Type /Page</c>, not <c>/Pages</c>).</summary>
    private static int Pages(string pdf) => Regex.Split(pdf, @"/Type\s*/Page[^s]").Length - 1;

    private static void QuarterlyData(Sheet s)
    {
        string[] quarters = ["Q1", "Q2", "Q3", "Q4"];
        s[0, 0].Value = T("Quarter");
        s[0, 1].Value = T("North");
        s[0, 2].Value = T("South");
        for (int r = 1; r <= 4; r++)
        {
            s[r, 0].Value = T(quarters[r - 1]);
            s[r, 1].Value = N(r * 100);
            s[r, 2].Value = N(r * 80);
        }
    }

    /// <summary>F24 a clustered column chart and a pie chart, reopened</summary>
    [Fact]
    public void CreateCharts()
    {
        using TempDir dir = new();
        string p = dir.Resolve("charts.xlsx");
        using (Workbook wb = new())
        {
            Sheet s = wb.Sheets.Add("Sales");
            QuarterlyData(s);
            Chart column = s.Charts.Add(ChartType.Column, new DrawingAnchor(6, 0, 20, 8));
            column.Title = "Quarterly Sales";
            column.Series.Add("North", "Sales!$B$2:$B$5", "Sales!$A$2:$A$5");
            column.Series.Add("South", "Sales!$C$2:$C$5", "Sales!$A$2:$A$5");
            column.Legend.Position = LegendPosition.Bottom;
            column.CategoryAxis.Title = "Quarter";
            column.ValueAxis.Title = "Units";
            Chart pie = s.Charts.Add(ChartType.Pie, new DrawingAnchor(6, 9, 20, 16));
            pie.Title = "North Share";
            pie.Series.Add("North", "Sales!$B$2:$B$5", "Sales!$A$2:$A$5");
            pie.Legend.Position = LegendPosition.Right;
            wb.Save(p);
        }
        using (Workbook wb = Workbook.Open(p))
        {
            Chart column = wb.Sheets["Sales"].Charts[0];
            Chart pie = wb.Sheets["Sales"].Charts[1];
            Assert.Equal(ChartType.Column, column.Type);
            Assert.Equal("Quarterly Sales", column.Title);
            Assert.Equal(2, column.Series.Count);
            Assert.Equal(ChartGrouping.Clustered, column.Grouping);
            Assert.Equal(LegendPosition.Bottom, column.Legend.Position);
            Assert.Equal(ChartType.Pie, pie.Type);
            Assert.Single(pie.Series);
        }
    }

    /// <summary>F25 change a line chart's title, fill, data labels and trendline, reopened</summary>
    [Fact]
    public void ModifyChart()
    {
        using TempDir dir = new();
        string p = dir.Resolve("line.xlsx");
        using (Workbook wb = new())
        {
            Sheet s = wb.Sheets.Add("Sales");
            QuarterlyData(s);
            Chart line = s.Charts.Add(ChartType.Line, new DrawingAnchor(6, 0, 20, 8));
            line.Title = "Sales";
            line.Series.Add("North", "Sales!$B$2:$B$5", "Sales!$A$2:$A$5");
            Chart c = s.Charts[line.Index];
            c.Title = "Sales (Revised)";
            c.Series[0].FillColor = Color.FromRgb(0x1F, 0x77, 0xB4);
            c.ShowDataLabels = true;
            c.Series[0].DataLabels = new DataLabels { ShowValue = true };
            c.Series[0].Trendlines.Add(new Trendline(TrendlineType.Linear));
            wb.Save(p);
        }
        using (Workbook wb = Workbook.Open(p))
        {
            Chart v = wb.Sheets["Sales"].Charts[0];
            Assert.Equal("Sales (Revised)", v.Title);
            Assert.Equal("FF1F77B4", v.Series[0].FillColor!.ToHex().Replace("#", ""));
            Assert.True(v.Series[0].DataLabels!.ShowValue);
            Assert.Single(v.Series[0].Trendlines);
        }
    }

    /// <summary>F26 one sparkline per data row, reopened</summary>
    [Fact]
    public void Sparklines()
    {
        using TempDir dir = new();
        string p = dir.Resolve("sparklines.xlsx");
        using (Workbook wb = new())
        {
            Sheet d = wb.Sheets.Add("Data");
            for (int r = 1; r <= 5; r++)
                for (int c = 1; c <= 12; c++)
                    d[r, c].Value = N((r * c) % 7);
            d.Sparklines.Add(SparklineType.Line, "Data!B2:M6", "N2:N6");
            wb.Save(p);
        }
        using (Workbook wb = Workbook.Open(p))
        {
            SparklineGroupCollection groups = wb.Sheets["Data"].Sparklines;
            Assert.Single(groups);
            Assert.Equal(SparklineType.Line, groups[0].Type);
            Assert.Equal(5, groups[0].Sparklines.Count);
        }
    }

    /// <summary>F27 a generated PNG anchored at a cell, reopened, then removed</summary>
    [Fact]
    public void Pictures()
    {
        // Java draws a 40 x 20 BufferedImage (left half filled) and writes it with ImageIO;
        // .NET has no portable imaging in the BCL, so the same PNG is encoded by hand.
        byte[] png = Png(40, 20, (x, _) => x < 20 ? 0xFFFFFF : 0x000000);
        using TempDir dir = new();
        string p = dir.Resolve("picture.xlsx");
        using (Workbook wb = new())
        {
            wb.Sheets.Add("S").Pictures.Add(png, PictureFormat.Png, new DrawingAnchor(1, 1, 6, 5));
            wb.Save(p);
        }
        using (Workbook wb = Workbook.Open(p))
        {
            Sheet s = wb.Sheets["S"];
            Assert.Single(s.Pictures);
            DrawingAnchor anchor = s.Pictures[0].Anchor;
            Assert.Equal([1, 1, 6, 5], new[] { anchor.From.Row, anchor.From.Column, anchor.To.Row, anchor.To.Column });
            Assert.Equal(PictureFormat.Png, s.Pictures[0].Format);
            // Java: removePicture(0) returns true; .NET: RemoveAt (throws when out of range)
            s.Pictures.RemoveAt(0);
            Assert.Empty(s.Pictures);
        }
    }

    /// <summary>F28 a filled rectangle with text and a text box, reopened</summary>
    [Fact]
    public void Shapes()
    {
        using TempDir dir = new();
        string p = dir.Resolve("shapes.xlsx");
        using (Workbook wb = new())
        {
            Sheet s = wb.Sheets.Add("S");
            Shape rect = s.Shapes.Add("rect", new DrawingAnchor(4, 7, 7, 11), "Review");
            rect.FillColor = Color.FromRgb(255, 204, 0);
            s.Shapes.AddTextBox(new DrawingAnchor(0, 7, 3, 11), "Quarterly sales");
            wb.Save(p);
        }
        using (Workbook wb = Workbook.Open(p))
        {
            Sheet s = wb.Sheets["S"];
            Assert.Equal(2, s.Shapes.Count);
            Assert.Equal("rect", s.Shapes[0].Geometry);
            Assert.Equal("Review", s.Shapes[0].Text);
            Assert.Equal("FFFFCC00", s.Shapes[0].FillColor!.ToHex().Replace("#", ""));
            Assert.True(s.Shapes[1].IsTextBox);
            Assert.Equal("Quarterly sales", s.Shapes[1].Text);
        }
    }

    /// <summary>F29 a pivot of region by product with the sum of amount</summary>
    [Fact]
    public void PivotTable()
    {
        using Workbook wb = new();
        Sheet d = wb.Sheets.Add("Data");
        string[][] rows =
        [
            ["Region", "Product", "Amount"], ["North", "Apple", "150"], ["North", "Banana", "80"],
            ["South", "Apple", "200"], ["South", "Banana", "90"],
        ];
        for (int r = 0; r < rows.Length; r++)
            for (int c = 0; c < 3; c++)
                d[r, c].Value = r > 0 && c == 2 ? N(double.Parse(rows[r][c])) : T(rows[r][c]);
        Sheet ps = wb.Sheets.Add("Pivot");
        Pivots.PivotTable pv = ps.PivotTables.Add("SalesPivot", "A1", "Data!$A$1:$C$5");
        pv.RowFields.Add("Region");
        pv.ColumnFields.Add("Product");
        pv.DataFields.Add("Amount", AggregateFunction.Sum);
        pv.Refresh();
        CellRange used = ps.UsedRange!;
        int grandRow = -1;
        for (int r = used.FirstRow; r <= used.LastRow; r++)
            if (T("Grand Total").Equals(ps[r, 0].Value))
                grandRow = r;
        Assert.True(grandRow > 0, "no Grand Total row");
        Assert.Equal(520.0, Number(ps[grandRow, used.LastColumn].Value));
    }

    /// <summary>F30 rename, move, copy, delete, hide, tab colour and active sheet, reopened</summary>
    [Fact]
    public void SheetManagement()
    {
        using TempDir dir = new();
        string p = dir.Resolve("sheets.xlsx");
        using (Workbook wb = new())
        {
            foreach (string name in new[] { "Data", "Notes", "Summary", "Temp", "Archive" })
                wb.Sheets.Add(name)[0, 0].Value = T(name);
            wb.Sheets["Notes"].Name = "Readme";
            wb.Sheets.Move(wb.Sheets["Summary"].Index, 0);
            wb.Sheets.AddCopy(wb.Sheets["Data"], "Data Copy");
            wb.Sheets.RemoveAt(wb.Sheets["Temp"].Index);
            wb.Sheets["Archive"].Visibility = Visibility.Hidden;
            wb.Sheets["Summary"].TabColor = Color.FromRgb(255, 128, 0);
            wb.ActiveSheet = wb.Sheets["Summary"];
            wb.Save(p);
        }
        using (Workbook wb = Workbook.Open(p))
        {
            List<string> names = wb.Sheets.Select(x => x.Name).Where(x => !x.StartsWith(Notice, StringComparison.Ordinal)).ToList();
            Assert.Equal(["Summary", "Data", "Readme", "Archive", "Data Copy"], names);
            Assert.Equal(Visibility.Hidden, wb.Sheets["Archive"].Visibility);
            Assert.Equal("FFFF8000", wb.Sheets["Summary"].TabColor!.ToHex().Replace("#", ""));
            Assert.Equal(T("Data"), wb.Sheets["Data Copy"][0, 0].Value);
        }
    }

    /// <summary>F31 format-only sheet protection, unlocked cells, structure and a protected range</summary>
    [Fact]
    public void Protection()
    {
        using TempDir dir = new();
        string p = dir.Resolve("protected.xlsx");
        using (Workbook wb = new())
        {
            Sheet s = wb.Sheets.Add("Data");
            s[1, 1].Style = new CellStyle { Locked = false };
            s.Protect("secret", new SheetProtection { FormatCells = true, FormatColumns = true, FormatRows = true });
            s.ProtectedRanges.Add("Totals", "D2:D5", "range-pw");
            wb.Protect("book-pw");
            wb.Save(p);
        }
        using (Workbook wb = Workbook.Open(p))
        {
            Sheet s = wb.Sheets["Data"];
            Assert.True(s.IsProtected);
            Assert.True(s.Protection!.FormatCells);
            Assert.False(s[1, 1].Style.Locked);
            Assert.True(s[0, 0].Style.Locked);
            Assert.Equal("Totals", s.ProtectedRanges[0].Name);
            Assert.True(wb.IsStructureLocked);
            Assert.True(s.Unprotect("secret"));
            Assert.False(s.IsProtected);
            Assert.True(wb.Unprotect("book-pw"));
        }
    }

    /// <summary>F32 an encrypted workbook reopens with its password; write protection verifies</summary>
    [Fact]
    public void EncryptionAndWriteProtection()
    {
        using TempDir dir = new();
        string encrypted = dir.Resolve("encrypted.xlsx");
        string writeProtected = dir.Resolve("write-protected.xlsx");
        using (Workbook wb = new())
        {
            wb.Sheets.Add("S")[0, 0].Value = T("Secret");
            // Java: saveEncrypted(path, pw); .NET: SaveOptions.Password
            wb.Save(encrypted, new SaveOptions { Password = "pw" });
            wb.SetWriteProtection("modify-pw", "Author", true);
            Assert.True(wb.VerifyWritePassword("modify-pw"));
            Assert.False(wb.VerifyWritePassword("wrong"));
            wb.Save(writeProtected);
        }
        using (Workbook wb = Workbook.Open(encrypted, new OpenOptions { Password = "pw" }))
        {
            Assert.Equal(T("Secret"), wb.Sheets["S"][0, 0].Value);
        }
        Assert.ThrowsAny<AxyraException>(() => Workbook.Open(encrypted).Dispose());
        using (Workbook wb = Workbook.Open(writeProtected))
        {
            Assert.True(wb.WriteProtection!.ReadOnlyRecommended);
            Assert.Equal("Author", wb.WriteProtection!.UserName);
            Assert.True(wb.VerifyWritePassword("modify-pw"));
        }
    }

    /// <summary>F33 page setup with a manual row page break, reopened</summary>
    [Fact]
    public void PageSetup()
    {
        using TempDir dir = new();
        string p = dir.Resolve("page.xlsx");
        using (Workbook wb = new())
        {
            Sheet s = wb.Sheets.Add("Report");
            for (int r = 0; r < 80; r++)
                s[r, 0].Value = N(r);
            Sheets.PageSetup page = s.PageSetup;
            page.Orientation = PageOrientation.Landscape;
            page.PaperSize = 9;
            page.Margins = new PageMargins(0.5, 0.5, 0.75, 0.75, 0.3, 0.3);
            page.PrintArea = "A1:F81";
            page.PrintTitleRows = "$1:$1";
            page.FitToPages(1, 0);
            page.Header = "&CQuarterly Sales Report";
            page.Footer = "&RPage &P of &N";
            page.AddRowBreak(40);
            wb.Save(p);
        }
        using (Workbook wb = Workbook.Open(p))
        {
            Sheets.PageSetup page = wb.Sheets["Report"].PageSetup;
            Assert.Equal(PageOrientation.Landscape, page.Orientation);
            Assert.Equal(9, page.PaperSize);
            Assert.Equal(0.75, page.Margins.Top, 1e-9);
            Assert.Equal("A1:F81", page.PrintArea);
            Assert.Equal("$1:$1", page.PrintTitleRows);
            // the evaluation build appends its notice to the header
            Assert.True(page.Header!.StartsWith("&CQuarterly Sales Report", StringComparison.Ordinal), page.Header);
            Assert.Equal([40], page.GetRowBreaks());
        }
    }

    /// <summary>F34 PDF export fitted to width, and a page range</summary>
    [Fact]
    public void PdfExport()
    {
        using TempDir dir = new();
        string all = dir.Resolve("out.pdf");
        string first = dir.Resolve("first.pdf");
        using (Workbook wb = new())
        {
            Sheet s = wb.Sheets.Add("Report");
            for (int r = 0; r < 200; r++)
            {
                s[r, 0].Value = T("Row " + r);
                s[r, 1].Value = N(r);
            }
            wb.SavePdf(all, new PdfOptions { SheetNames = ["Report"], FitTo = PdfFit.Width });
            wb.SavePdf(first, new PdfOptions { SheetNames = ["Report"], PageRange = "1" });
        }
        string pdf = Encoding.Latin1.GetString(File.ReadAllBytes(all));
        Assert.StartsWith("%PDF", pdf, StringComparison.Ordinal);
        Assert.True(Pages(pdf) > 1, "expected several pages, got " + Pages(pdf));
        string one = Encoding.Latin1.GetString(File.ReadAllBytes(first));
        Assert.Equal(1, Pages(one));
    }

    /// <summary>F35 render a sheet to PNG at 2x and to SVG</summary>
    [Fact]
    public void RenderImages()
    {
        using Workbook wb = new();
        Sheet s = wb.Sheets.Add("Report");
        for (int r = 0; r < 10; r++)
            s[r, 0].Value = N(r);
        byte[] png = s.RenderImage(ImageFormat.Png, new ImageOptions { Scale = 2.0 });
        byte[] svg = s.RenderImage(ImageFormat.Svg, new ImageOptions());
        Assert.Equal(new byte[] { 0x89, (byte)'P', (byte)'N', (byte)'G' }, png[..4]);
        int width = (png[16] << 24) | (png[17] << 16) | (png[18] << 8) | png[19];
        Assert.True(width > 1500 && width < 1700, "A4 at 2x is about 1588 px wide, was " + width);
        Assert.Contains("<svg", Encoding.UTF8.GetString(svg));
        Assert.Equal(1, s.GetPageCount());
    }

    /// <summary>A minimal RGB PNG encoder (the stand-in for Java's ImageIO).</summary>
    private static byte[] Png(int width, int height, Func<int, int, int> rgb)
    {
        using MemoryStream raw = new();
        for (int y = 0; y < height; y++)
        {
            raw.WriteByte(0); // filter: none
            for (int x = 0; x < width; x++)
            {
                int c = rgb(x, y);
                raw.WriteByte((byte)(c >> 16));
                raw.WriteByte((byte)(c >> 8));
                raw.WriteByte((byte)c);
            }
        }
        using MemoryStream idat = new();
        using (ZLibStream z = new(idat, CompressionLevel.Optimal, leaveOpen: true))
            raw.WriteTo(z);

        using MemoryStream png = new();
        png.Write([0x89, (byte)'P', (byte)'N', (byte)'G', 0x0D, 0x0A, 0x1A, 0x0A]);
        byte[] ihdr = new byte[13];
        WriteBigEndian(ihdr, 0, width);
        WriteBigEndian(ihdr, 4, height);
        ihdr[8] = 8; // bit depth
        ihdr[9] = 2; // colour type: RGB
        Chunk(png, "IHDR", ihdr);
        Chunk(png, "IDAT", idat.ToArray());
        Chunk(png, "IEND", []);
        return png.ToArray();
    }

    private static void Chunk(Stream output, string type, byte[] data)
    {
        byte[] header = new byte[8];
        WriteBigEndian(header, 0, data.Length);
        Encoding.ASCII.GetBytes(type, 0, 4, header, 4);
        output.Write(header);
        output.Write(data);
        uint crc = 0xFFFFFFFF;
        foreach (byte b in header.AsSpan(4).ToArray().Concat(data))
        {
            crc ^= b;
            for (int k = 0; k < 8; k++)
                crc = (crc & 1) != 0 ? (crc >> 1) ^ 0xEDB88320 : crc >> 1;
        }
        byte[] tail = new byte[4];
        WriteBigEndian(tail, 0, (int)~crc);
        output.Write(tail);
    }

    private static void WriteBigEndian(byte[] buffer, int offset, int value)
    {
        buffer[offset] = (byte)(value >> 24);
        buffer[offset + 1] = (byte)(value >> 16);
        buffer[offset + 2] = (byte)(value >> 8);
        buffer[offset + 3] = (byte)value;
    }
}
