namespace Axyra.Sheets.Examples.Tests;

public class GuideExamplesIntegrationTest
{
    private static readonly IReadOnlyList<string> ExpectedGuides =
    [
        "guides/convert-excel-to-pdf",
        "guides/create-excel-file",
        "guides/generate-report",
        "guides/read-excel-file",
        "guides/write-data-to-excel",
    ];

    private static readonly IReadOnlyDictionary<string, IReadOnlyList<string>> ExpectedOutputs =
        new Dictionary<string, IReadOnlyList<string>>
        {
            ["guides/convert-excel-to-pdf"] = ["pdf-source.xlsx", "report.pdf"],
            ["guides/create-excel-file"] = ["q3-sales.xlsx"],
            ["guides/generate-report"] = ["sales-template.xlsx", "sales-Q3-2026.xlsx", "sales-Q3-2026.pdf"],
            ["guides/read-excel-file"] = ["read-source.xlsx"],
            ["guides/write-data-to-excel"] = ["data-export.xlsx", "streaming-export.xlsx"],
        };

    [Fact]
    public void RunsEveryGuideExample()
    {
        using TempDir output = new();
        List<IExample> guides = ExampleRunner.Examples
            .Where(example => example.Name.StartsWith("guides/", StringComparison.Ordinal))
            .ToList();

        Assert.Equal(ExpectedGuides, guides.Select(e => e.Name).Order(StringComparer.Ordinal).ToList());

        foreach (IExample guide in guides)
        {
            string guideOutput = Directory.CreateDirectory(
                output.Resolve(guide.Name["guides/".Length..])).FullName;
            guide.Run(guideOutput);
            foreach (string fileName in ExpectedOutputs[guide.Name])
            {
                string file = Path.Combine(guideOutput, fileName);
                Assert.True(IsNonEmptyFile(file), guide.Name + " did not produce " + fileName);
            }
        }
    }

    private static bool IsNonEmptyFile(string path)
    {
        try
        {
            return File.Exists(path) && new FileInfo(path).Length > 0;
        }
        catch (Exception)
        {
            return false;
        }
    }
}
