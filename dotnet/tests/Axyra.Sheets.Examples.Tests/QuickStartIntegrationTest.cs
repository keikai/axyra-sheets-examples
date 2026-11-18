using Axyra.Sheets.Examples.Tutorial;

namespace Axyra.Sheets.Examples.Tests;

public class QuickStartIntegrationTest
{
    [Fact]
    public void WritesARealWorkbook()
    {
        using TempDir output = new();
        new QuickStartExample().Run(output.Path);
        string workbookPath = output.Resolve("hello.xlsx");
        Assert.True(new FileInfo(workbookPath).Length > 0);
        using Workbook workbook = Workbook.Open(workbookPath);
        Assert.Equal("Widget", ((CellValue.Text)workbook.Sheets["Sheet1"][1, 0].Value).Value);
        Assert.Equal(51.0, ((CellValue.Number)workbook.Sheets["Sheet1"][1, 3].Value).Value);
    }
}
