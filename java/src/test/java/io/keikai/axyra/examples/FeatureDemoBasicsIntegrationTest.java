package io.keikai.axyra.examples;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.nio.file.Path;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/** Feature demos for workbook manipulation, reading and writing, and formatting. */
class FeatureDemoBasicsIntegrationTest {
    static final Map<String, List<String>> EXPECTED_OUTPUTS = new LinkedHashMap<>();

    static {
        EXPECTED_OUTPUTS.put("feature-demo/create-workbook", List.of(
                "create-workbook-sparse.xlsx", "create-workbook-columnar.xlsx",
                "create-workbook-defaults.xlsx"));
        EXPECTED_OUTPUTS.put("feature-demo/open-workbook", List.of(
                "open-workbook-from-stream.xlsx", "open-workbook-from-bytes.xlsx",
                "open-workbook-inspect.xlsx", "open-workbook-recalculated.xlsx"));
        EXPECTED_OUTPUTS.put("feature-demo/save-export", List.of(
                "quarterly-report.xlsx", "quarterly-report.csv",
                "quarterly-report.json", "quarterly-report.html"));
        EXPECTED_OUTPUTS.put("feature-demo/manage-worksheets", List.of(
                "manage-worksheets-create-insert.xlsx", "manage-worksheets-rename-reorder.xlsx",
                "manage-worksheets-copy.xlsx", "manage-worksheets-delete.xlsx",
                "manage-worksheets-tab-colours.xlsx"));
        EXPECTED_OUTPUTS.put("feature-demo/cell-read-write", List.of(
                "cell-read-write-types.xlsx", "cell-read-write-read-back.xlsx",
                "cell-read-write-dates.xlsx"));
        EXPECTED_OUTPUTS.put("feature-demo/bulk-range", List.of(
                "bulk-range-mixed-block.xlsx", "bulk-range-numeric.xlsx",
                "bulk-range-read-back.xlsx"));
        EXPECTED_OUTPUTS.put("feature-demo/formulas-recalculation", List.of(
                "formulas-calculate.xlsx", "formulas-before-after.xlsx", "formulas-ad-hoc.xlsx",
                "formulas-dynamic-arrays.xlsx", "formulas-manual-mode.xlsx"));
        EXPECTED_OUTPUTS.put("feature-demo/import-objects", List.of(
                "import-objects-beans.xlsx", "import-objects-chosen-columns.xlsx",
                "import-objects-maps.xlsx", "import-objects-total.xlsx"));
        EXPECTED_OUTPUTS.put("feature-demo/cell-styling", List.of(
                "cell-styling-fonts.xlsx", "cell-styling-fills.xlsx", "cell-styling-borders.xlsx",
                "cell-styling-alignment.xlsx", "cell-styling-report.xlsx"));
        EXPECTED_OUTPUTS.put("feature-demo/number-formats", List.of(
                "number-formats-numbers.xlsx", "number-formats-dates.xlsx",
                "number-formats-locales.xlsx", "number-formats-stateless.xlsx"));
        EXPECTED_OUTPUTS.put("feature-demo/rows-columns", List.of(
                "rows-columns-sizes.xlsx", "rows-columns-autofit.xlsx",
                "rows-columns-hidden.xlsx", "rows-columns-insert.xlsx"));
        EXPECTED_OUTPUTS.put("feature-demo/merge-freeze", List.of(
                "merge-freeze-title-band.xlsx", "merge-freeze-auto-merge.xlsx",
                "merge-freeze-list-merges.xlsx", "merge-freeze-frozen.xlsx"));
    }

    @Test
    void runsEveryBasicsFeatureDemo(@TempDir Path output) throws Exception {
        FeatureDemoTestSupport.runAndCheck(EXPECTED_OUTPUTS, output);
    }
}
