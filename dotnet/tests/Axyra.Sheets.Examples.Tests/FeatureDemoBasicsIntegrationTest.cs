namespace Axyra.Sheets.Examples.Tests;

/// <summary>Feature demos for workbook manipulation, reading and writing, and formatting.</summary>
public class FeatureDemoBasicsIntegrationTest
{
    internal static readonly IReadOnlyDictionary<string, string[]> ExpectedOutputs =
        new Dictionary<string, string[]>
        {
            ["feature-demo/create-workbook"] =
            [
                "create-workbook-sparse.xlsx", "create-workbook-columnar.xlsx",
                "create-workbook-defaults.xlsx",
            ],
            ["feature-demo/open-workbook"] =
            [
                "open-workbook-from-stream.xlsx", "open-workbook-from-bytes.xlsx",
                "open-workbook-inspect.xlsx", "open-workbook-recalculated.xlsx",
            ],
            ["feature-demo/save-export"] =
            [
                "quarterly-report.xlsx", "quarterly-report.csv",
                "quarterly-report.json", "quarterly-report.html",
            ],
            ["feature-demo/manage-worksheets"] =
            [
                "manage-worksheets-create-insert.xlsx", "manage-worksheets-rename-reorder.xlsx",
                "manage-worksheets-copy.xlsx", "manage-worksheets-delete.xlsx",
                "manage-worksheets-tab-colours.xlsx",
            ],
            ["feature-demo/cell-read-write"] =
            [
                "cell-read-write-types.xlsx", "cell-read-write-read-back.xlsx",
                "cell-read-write-dates.xlsx",
            ],
            ["feature-demo/bulk-range"] =
            [
                "bulk-range-mixed-block.xlsx", "bulk-range-numeric.xlsx",
                "bulk-range-read-back.xlsx",
            ],
            ["feature-demo/formulas-recalculation"] =
            [
                "formulas-calculate.xlsx", "formulas-before-after.xlsx", "formulas-ad-hoc.xlsx",
                "formulas-dynamic-arrays.xlsx", "formulas-manual-mode.xlsx",
            ],
            ["feature-demo/import-objects"] =
            [
                "import-objects-beans.xlsx", "import-objects-chosen-columns.xlsx",
                "import-objects-maps.xlsx", "import-objects-total.xlsx",
            ],
            ["feature-demo/cell-styling"] =
            [
                "cell-styling-fonts.xlsx", "cell-styling-fills.xlsx", "cell-styling-borders.xlsx",
                "cell-styling-alignment.xlsx", "cell-styling-report.xlsx",
            ],
            ["feature-demo/number-formats"] =
            [
                "number-formats-numbers.xlsx", "number-formats-dates.xlsx",
                "number-formats-locales.xlsx", "number-formats-stateless.xlsx",
            ],
            ["feature-demo/rows-columns"] =
            [
                "rows-columns-sizes.xlsx", "rows-columns-autofit.xlsx",
                "rows-columns-hidden.xlsx", "rows-columns-insert.xlsx",
            ],
            ["feature-demo/merge-freeze"] =
            [
                "merge-freeze-title-band.xlsx", "merge-freeze-auto-merge.xlsx",
                "merge-freeze-list-merges.xlsx", "merge-freeze-frozen.xlsx",
            ],
        };

    [Fact]
    public void RunsEveryBasicsFeatureDemo()
    {
        using TempDir output = new();
        FeatureDemoTestSupport.RunAndCheck(ExpectedOutputs, output.Path);
    }
}
