using Axyra.Sheets.Content;
using Axyra.Sheets.Style;

namespace Axyra.Sheets.Examples.DevRef;

public sealed class ContentObjectsExample : IExample
{
    public string Name => "dev-ref/content-objects";

    public string Description => "Create a table, validation, conditional formatting, comments, and links";

    public void Run(string outputDirectory)
    {
        using Workbook workbook = new Workbook();
        Sheet sheet = workbook.Sheets.Add("Content");
        sheet["A1:D4"].SetValues(new CellValue[,]
        {
            { "Region", "Q1", "Q2", "Total" },
            { "APAC", 1200, 1450, 2650 },
            { "EMEA", 980, 1100, 2080 },
            { "Americas", 1610, 1720, 3330 },
        });
        sheet.Tables.Add("Sales", "A1:D4", "Region", "Q1", "Q2", "Total");

        CellRange totals = sheet["D2:D4"];
        sheet.ConditionalFormats.Add(totals, new CellValueRule(ComparisonOperator.GreaterThan, "2500")
        {
            Style = new ConditionalStyle
            {
                FillColor = Color.Parse("#FFF3CD"),
                FontColor = Color.Parse("#8A6D3B"),
                Bold = true,
            },
            Priority = 1,
        });
        sheet.Validations.Add("A2:A20", Validation.List("APAC", "EMEA", "Americas"));

        Cell note = sheet[1, 3];
        note.Comment = new Comment("Finance", "Adjusted after audit");
        note.Hyperlink = Hyperlink.Url("https://keikai.io", "Keikai");
        sheet[5, 0].RichText = new RichText(
            new TextRun("Total: "),
            new TextRun("reviewed") { Bold = true, Color = Color.Parse("#C0392B") });

        sheet.Protect("pa55w0rd");
        sheet.ProtectedRanges.Add("Editors", "B2:D4", "range-pw");
        workbook.Protect("pa55w0rd");   // locks the structure
        workbook.Save(Path.Combine(outputDirectory, "content-objects.xlsx"));
    }
}
