namespace Axyra.Sheets.Examples.Tests;

/// <summary>Feature demos for data manipulation, and structure and output.</summary>
public class FeatureDemoDataAndOutputIntegrationTest
{
    // A Dictionary that is never removed from enumerates in insertion order, as Java's
    // LinkedHashMap does (OrderedDictionary<,> needs .NET 9).
    internal static readonly IReadOnlyDictionary<string, string[]> ExpectedOutputs =
        new Dictionary<string, string[]>
        {
            ["feature-demo/find-replace"] =
            [
                "find-replace-highlight.xlsx", "find-replace-options.xlsx",
                "find-replace-replace-all.xlsx", "find-replace-regex.xlsx",
            ],
            ["feature-demo/sort-filter"] =
            [
                "sort-filter-unsorted.xlsx", "sort-filter-sort-one.xlsx", "sort-filter-sort-two.xlsx",
                "sort-filter-filter-value.xlsx", "sort-filter-filter-comparison.xlsx",
            ],
            ["feature-demo/data-validation"] =
            [
                "data-validation-dropdown.xlsx", "data-validation-numeric.xlsx",
                "data-validation-prompts.xlsx", "data-validation-custom.xlsx",
            ],
            ["feature-demo/conditional-format"] =
            [
                "conditional-format-threshold.xlsx", "conditional-format-data-bars.xlsx",
                "conditional-format-color-scale.xlsx", "conditional-format-icon-set.xlsx",
                "conditional-format-row-rule.xlsx",
            ],
            ["feature-demo/create-table"] =
            [
                "create-table-named.xlsx", "create-table-inferred.xlsx",
                "create-table-structured.xlsx", "create-table-totals.xlsx",
            ],
            ["feature-demo/insert-chart"] =
            [
                "insert-chart-column.xlsx", "insert-chart-line.xlsx", "insert-chart-area.xlsx",
                "insert-chart-pie.xlsx", "insert-chart-two-charts.xlsx",
            ],
            ["feature-demo/export-pdf"] =
            [
                "export-pdf-report.xlsx", "report.pdf", "report-sheet1.pdf",
                "report-fit-width.pdf", "export-pdf-page-count.xlsx",
            ],
            ["feature-demo/render-image"] =
            [
                "render-image-dashboard.xlsx", "sheet.png", "sheet-2x.png",
                "sheet-no-gridlines.png", "sheet.svg",
            ],
        };

    [Fact]
    public void RunsEveryDataAndOutputFeatureDemo()
    {
        using TempDir output = new();
        FeatureDemoTestSupport.RunAndCheck(ExpectedOutputs, output.Path);
    }
}
