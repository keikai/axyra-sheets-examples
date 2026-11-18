package io.keikai.axyra.examples;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.nio.file.Path;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/** Feature demos for data manipulation, and structure and output. */
class FeatureDemoDataAndOutputIntegrationTest {
    static final Map<String, List<String>> EXPECTED_OUTPUTS = new LinkedHashMap<>();

    static {
        EXPECTED_OUTPUTS.put("feature-demo/find-replace", List.of(
                "find-replace-highlight.xlsx", "find-replace-options.xlsx",
                "find-replace-replace-all.xlsx", "find-replace-regex.xlsx"));
        EXPECTED_OUTPUTS.put("feature-demo/sort-filter", List.of(
                "sort-filter-unsorted.xlsx", "sort-filter-sort-one.xlsx", "sort-filter-sort-two.xlsx",
                "sort-filter-filter-value.xlsx", "sort-filter-filter-comparison.xlsx"));
        EXPECTED_OUTPUTS.put("feature-demo/data-validation", List.of(
                "data-validation-dropdown.xlsx", "data-validation-numeric.xlsx",
                "data-validation-prompts.xlsx", "data-validation-custom.xlsx"));
        EXPECTED_OUTPUTS.put("feature-demo/conditional-format", List.of(
                "conditional-format-threshold.xlsx", "conditional-format-data-bars.xlsx",
                "conditional-format-color-scale.xlsx", "conditional-format-icon-set.xlsx",
                "conditional-format-row-rule.xlsx"));
        EXPECTED_OUTPUTS.put("feature-demo/create-table", List.of(
                "create-table-named.xlsx", "create-table-inferred.xlsx",
                "create-table-structured.xlsx", "create-table-totals.xlsx"));
        EXPECTED_OUTPUTS.put("feature-demo/insert-chart", List.of(
                "insert-chart-column.xlsx", "insert-chart-line.xlsx", "insert-chart-area.xlsx",
                "insert-chart-pie.xlsx", "insert-chart-two-charts.xlsx"));
        EXPECTED_OUTPUTS.put("feature-demo/export-pdf", List.of(
                "export-pdf-report.xlsx", "report.pdf", "report-sheet1.pdf",
                "report-fit-width.pdf", "export-pdf-page-count.xlsx"));
        EXPECTED_OUTPUTS.put("feature-demo/render-image", List.of(
                "render-image-dashboard.xlsx", "sheet.png", "sheet-2x.png",
                "sheet-no-gridlines.png", "sheet.svg"));
    }

    @Test
    void runsEveryDataAndOutputFeatureDemo(@TempDir Path output) throws Exception {
        FeatureDemoTestSupport.runAndCheck(EXPECTED_OUTPUTS, output);
    }
}
