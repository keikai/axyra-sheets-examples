using Axyra.Sheets.Content;
using Axyra.Sheets.Style;
using Axyra.Sheets.Tables;
using static Axyra.Sheets.Examples.Tests.Regression.IssueTestSupport;

namespace Axyra.Sheets.Examples.Tests.Regression;

/// <summary>
/// Features F14-F23 of the axyra-sheets skill test (merging, panes, rows and columns, outlines,
/// sorting, filtering, find and replace, validation, conditional formatting, tables).
/// </summary>
public class FeatureLayoutTest
{
    /// <summary>F14 merge A1:D1, list merged regions, unmerge</summary>
    [Fact]
    public void MergeCells()
    {
        using Workbook wb = new();
        Sheet s = wb.Sheets.Add("S");
        s[0, 0].Value = T("Title");
        s["A1:D1"].Merge();
        Assert.Single(s.MergedRegions);
        Assert.Equal(3, s.MergedRegions[0].LastColumn);
        s["A1:D1"].Unmerge();
        Assert.Empty(s.MergedRegions);
    }

    /// <summary>F15 frozen and split panes survive save and reopen</summary>
    [Fact]
    public void FreezeAndSplit()
    {
        using TempDir dir = new();
        string p = dir.Resolve("freeze.xlsx");
        using (Workbook wb = new())
        {
            wb.Sheets.Add("Frozen").FreezePanes(1, 1);
            wb.Sheets.Add("Split").SplitPanes(5, 3);
            wb.Save(p);
        }
        using (Workbook wb = Workbook.Open(p))
        {
            Assert.Equal(1, wb.Sheets["Frozen"].FrozenRows);
            Assert.Equal(1, wb.Sheets["Frozen"].FrozenColumns);
            Assert.Equal(5, wb.Sheets["Split"].SplitRow);
            Assert.Equal(3, wb.Sheets["Split"].SplitColumn);
        }
    }

    /// <summary>F16 insert and delete rows and columns, widths, heights, hidden and auto-fit</summary>
    [Fact]
    public void RowsAndColumns()
    {
        using TempDir dir = new();
        string p = dir.Resolve("rows.xlsx");
        using (Workbook wb = new())
        {
            Sheet s = wb.Sheets.Add("S");
            s[0, 0].Value = T("Name");
            s[1, 0].Value = N(1);
            s[2, 0].Value = N(2);
            s.Rows.Insert(1, 2);
            s.Columns.Insert(0, 1);
            s.Rows.Delete(1, 1);
            Assert.Equal(N(2), s[3, 1].Value);
            s.Columns[0].Width = 18;
            s.Rows[0].Height = 30;
            s.Columns[2].Hidden = true;
            s[0, 3].Value = T("a fairly long piece of text for auto-fit");
            s.Columns[3].AutoFit();
            Assert.True(s.Columns[3].Width > 20);
            wb.Save(p);
        }
        using (Workbook wb = Workbook.Open(p))
        {
            Sheet s = wb.Sheets["S"];
            Assert.Equal(18.0, s.Columns[0].Width, 0.01);
            Assert.Equal(30.0, s.Rows[0].Height, 0.01);
            Assert.True(s.Columns[2].Hidden);
        }
    }

    /// <summary>F17 grouped and collapsed rows, grouped columns, subtotals by category</summary>
    [Fact]
    public void GroupingAndSubtotals()
    {
        using TempDir dir = new();
        string p = dir.Resolve("outline.xlsx");
        using (Workbook wb = new())
        {
            Sheet outline = wb.Sheets.Add("Outline");
            for (int r = 0; r < 7; r++)
                outline[r, 0].Value = N(r);
            outline.Rows.Group(1, 4);
            outline.Rows.SetGroupCollapsed(1, 4, true);
            outline.Columns.Group(1, 2);
            Sheet sales = wb.Sheets.Add("Sales");
            string[][] rows =
            [
                ["Category", "Item", "Amount"], ["Dairy", "Milk", "60"], ["Dairy", "Cheese", "150"],
                ["Fruit", "Apple", "120"], ["Fruit", "Banana", "80"],
            ];
            for (int r = 0; r < rows.Length; r++)
                for (int c = 0; c < 3; c++)
                    sales[r, c].Value = r > 0 && c == 2 ? N(double.Parse(rows[r][c])) : T(rows[r][c]);
            // Java: sales.subtotal(firstRow 0, lastRow 4, groupBy 0, hasHeader true, SUM, column 2)
            Assert.True(sales["A1:C5"].Subtotal(0, SubtotalFunction.Sum, [2], hasHeaders: true) > 0);
            wb.Save(p);
        }
        using (Workbook wb = Workbook.Open(p))
        {
            Sheet outline = wb.Sheets["Outline"];
            Assert.Equal(1, outline.Rows[1].OutlineLevel);
            Assert.True(outline.Rows[1].Hidden);
            Assert.Equal(0, outline.Rows[5].OutlineLevel);
            Sheet sales = wb.Sheets["Sales"];
            Assert.Equal("SUBTOTAL(9,C2:C3)", sales[3, 2].Formula);
            Assert.Equal(210.0, Number(sales[3, 2].Value));
        }
    }

    /// <summary>F18 sort rows by two keys and sort columns left to right</summary>
    [Fact]
    public void Sorting()
    {
        using Workbook wb = new();
        Sheet s = wb.Sheets.Add("S");
        (string Name, int Score)[] data = [("Eve", 80), ("Bob", 95), ("Dan", 80), ("Alice", 95), ("Carl", 70)];
        for (int r = 0; r < data.Length; r++)
        {
            s[r, 0].Value = T(data[r].Name);
            s[r, 1].Value = N(data[r].Score);
        }
        s["A1:B5"].Sort(new SortKey(1, Ascending: false), new SortKey(0));
        Assert.Equal([T("Alice"), T("Bob"), T("Dan"), T("Eve"), T("Carl")],
            new[] { s[0, 0].Value, s[1, 0].Value, s[2, 0].Value, s[3, 0].Value, s[4, 0].Value });

        Sheet h = wb.Sheets.Add("H");
        string[] names = ["Delta", "Alpha", "Charlie"];
        for (int c = 0; c < names.Length; c++)
        {
            h[0, c].Value = T(names[c]);
            h[1, c].Value = N(c);
        }
        h["A1:C2"].SortLeftToRight(new SortKey(0));
        Assert.Equal(T("Alpha"), h[0, 0].Value);
        Assert.Equal(T("Delta"), h[0, 2].Value);
    }

    /// <summary>F19 AutoFilter by values and a greater-than condition, saved and reopened</summary>
    [Fact]
    public void AutoFilter()
    {
        using TempDir dir = new();
        string p = dir.Resolve("filter.xlsx");
        using (Workbook wb = new())
        {
            Sheet s = wb.Sheets.Add("S");
            object[][] rows =
            [
                ["Region", "Amount"], ["North", 120], ["South", 80], ["East", 200], ["North", 90], ["South", 140],
            ];
            for (int r = 0; r < rows.Length; r++)
            {
                s[r, 0].Value = T((string)rows[r][0]);
                s[r, 1].Value = r == 0 ? T((string)rows[r][1]) : N((int)rows[r][1]);
            }
            s.AutoFilter.Range = s["A1:B6"];
            s.AutoFilter[0].FilterByValues("North", "South");
            s.AutoFilter[1].FilterByCondition(FilterOperator.GreaterThan, "100");
            s.AutoFilter.Apply();
            wb.Save(p);
        }
        using (Workbook wb = Workbook.Open(p))
        {
            Sheet s = wb.Sheets["S"];
            Assert.False(s.Rows[1].Hidden);   // North 120
            Assert.True(s.Rows[2].Hidden);    // South 80
            Assert.True(s.Rows[3].Hidden);    // East
            Assert.True(s.Rows[4].Hidden);    // North 90
            Assert.False(s.Rows[5].Hidden);   // South 140
        }
    }

    /// <summary>F20 find, replace and regex replace</summary>
    [Fact]
    public void FindAndReplace()
    {
        using Workbook wb = new();
        Sheet s = wb.Sheets.Add("S");
        s[0, 0].Value = T("Apple");
        s[1, 0].Value = T("Fresh apple");
        s[2, 0].Value = T("Banana");
        s[0, 1].Value = T("ID-1001");
        s[1, 1].Value = T("ID-2002");
        Assert.Equal(2, s.FindAll("apple").Count);
        Assert.Equal(2, s.ReplaceAll("apple", "pear"));
        Assert.Equal(T("Fresh pear"), s[1, 0].Value);
        // Java: replaceAllRegex; .NET: ReplaceAll with FindOptions.UseRegex
        Assert.Equal(2, s.ReplaceAll(@"ID-\d+", "ID-X", new FindOptions { UseRegex = true }));
        Assert.Equal(T("ID-X"), s[1, 1].Value);
    }

    /// <summary>F21 a drop-down list and a whole-number rule with an error message, reopened</summary>
    [Fact]
    public void DataValidation()
    {
        using TempDir dir = new();
        string p = dir.Resolve("validation.xlsx");
        using (Workbook wb = new())
        {
            Sheet s = wb.Sheets.Add("S");
            s.Validations.Add("B2:B100", Validation.List("Open", "In Progress", "Closed"));
            s.Validations.Add("C2:C100", Validation.WholeNumber(ValidationOperator.Between, "1", "100") with
            {
                ErrorTitle = "Invalid quantity",
                ErrorMessage = "Enter 1 to 100.",
            });
            wb.Save(p);
        }
        using (Workbook wb = Workbook.Open(p))
        {
            List<Validation> rules = wb.Sheets["S"].Validations.Select(v => v.Rule).ToList();
            Assert.Equal(2, rules.Count);
            Assert.Contains(rules, v => v.Type == ValidationType.List);
            Validation whole = rules.First(v => v.Type == ValidationType.WholeNumber);
            Assert.Equal("100", whole.Formula2);
            Assert.Equal("Invalid quantity", whole.ErrorTitle);
        }
    }

    /// <summary>F22 cell-value rule, three-colour scale and data bar, reopened</summary>
    [Fact]
    public void ConditionalFormatting()
    {
        using TempDir dir = new();
        string p = dir.Resolve("cf.xlsx");
        using (Workbook wb = new())
        {
            Sheet s = wb.Sheets.Add("S");
            for (int r = 0; r < 10; r++)
                s[r, 0].Value = N(r * 25);
            s.ConditionalFormats.Add("A1:A10", new CellValueRule(ComparisonOperator.GreaterThan, "100")
            {
                Style = new ConditionalStyle { FillColor = Color.Red },
            });
            s.ConditionalFormats.Add("B1:B10", ColorScaleRule.ThreeColor(Color.Red, Color.Yellow, Color.Green));
            s.ConditionalFormats.Add("C1:C10", new DataBarRule(Color.Blue));
            wb.Save(p);
        }
        using (Workbook wb = Workbook.Open(p))
        {
            // Java compares the type tokens "cellIs", "colorScale", "dataBar"; .NET has one record per rule type
            List<Type> types = wb.Sheets["S"].ConditionalFormats.Select(c => c.Rule.GetType()).ToList();
            Assert.Equal([typeof(CellValueRule), typeof(ColorScaleRule), typeof(DataBarRule)], types);
        }
    }

    /// <summary>F23 a styled table with a totals row, resized, then converted to a range</summary>
    [Fact]
    public void Tables()
    {
        using Workbook wb = new();
        Sheet s = wb.Sheets.Add("S");
        string[] header = ["Region", "Product", "Amount"];
        for (int c = 0; c < 3; c++)
            s[0, c].Value = T(header[c]);
        for (int r = 1; r <= 6; r++)
        {
            s[r, 0].Value = T("R" + r);
            s[r, 1].Value = T("P" + r);
            s[r, 2].Value = N(r * 100);
        }
        // Java's setTableStyle / resizeTable / setTableTotalsRow / setTableColumnTotals return true;
        // the .NET counterparts are property setters on the live Table (they throw on failure).
        Table table = s.Tables.Add("SalesTable", s["A1:C5"], header);
        table.Style = new TableStyle("TableStyleMedium2");
        table.Range = s["A1:C7"];
        table.ShowTotalsRow = true;
        table.Columns["Amount"].TotalsFunction = TotalsRowFunction.Sum;
        Assert.Equal("SalesTable", s.Tables[0].Name);
        Assert.Equal(header, s.Tables[0].Columns.Select(c => c.Name).ToArray());
        s.Tables["SalesTable"].ConvertToRange();
        Assert.Empty(s.Tables);
        Assert.Equal(T("Region"), s[0, 0].Value);
    }
}
