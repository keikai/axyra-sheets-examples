package io.keikai.axyra.examples.featuredemo;

import io.keikai.axyra.examples.Example;
import io.keikai.axyra.sheets.CellValue;
import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.Workbook;
import io.keikai.axyra.sheets.content.ConditionalFormat;
import io.keikai.axyra.sheets.style.CellStyle;
import io.keikai.axyra.sheets.style.Color;

import java.nio.file.Path;

/**
 * Conditional formatting — styling driven by the values themselves.
 *
 * <p>A rule is attached to a range and re-evaluated by the spreadsheet
 * application, so editing a cell restyles it without any Java running. Rules
 * stack: several on one range apply in order, which is how a heat map and a
 * threshold flag coexist.
 */
public final class ConditionalFormatExample implements Example {
    private static final String CONTENT = "conditionalFormat-figures-content.xlsx";
    private static final int ROWS = 18;

    @Override
    public String name() {
        return "feature-demo/conditional-format";
    }

    @Override
    public String description() {
        return "Add threshold, data bar, colour scale, icon set and formula rules";
    }

    @Override
    public void run(Path outputDirectory) {
        try (Workbook wb = threshold()) {
            wb.save(outputDirectory.resolve("conditional-format-threshold.xlsx"));
        }
        try (Workbook wb = dataBars()) {
            wb.save(outputDirectory.resolve("conditional-format-data-bars.xlsx"));
        }
        try (Workbook wb = colorScale()) {
            wb.save(outputDirectory.resolve("conditional-format-color-scale.xlsx"));
        }
        try (Workbook wb = iconSet()) {
            wb.save(outputDirectory.resolve("conditional-format-icon-set.xlsx"));
        }
        try (Workbook wb = rowRule()) {
            wb.save(outputDirectory.resolve("conditional-format-row-rule.xlsx"));
        }
        System.out.println("Wrote five conditional-formatting workbooks to " + outputDirectory);
    }

    /** Two stacked rules on one column: over target, and under water. */
    private static Workbook threshold() {
        Workbook wb = figures();
        Sheet sheet = wb.sheet(0);

        sheet.range(1, 3, ROWS, 3).addConditionalFormat(ConditionalFormat.builder()
                .cellIs("greaterThan", "3000")
                .fillColor(Color.rgb(0xC6, 0xEF, 0xCE))
                .fontColor(Color.rgb(0x00, 0x61, 0x00))
                .bold()
                .build());

        sheet.range(1, 4, ROWS, 4).addConditionalFormat(ConditionalFormat.builder()
                .cellIs("lessThan", "0")
                .fillColor(Color.rgb(0xFF, 0xC7, 0xCE))
                .fontColor(Color.rgb(0x9C, 0x00, 0x06))
                .build());

        note(sheet, "Revenue over 3,000 is green; a negative variance is red. "
                + "Edit a cell and the styling follows.");
        return wb;
    }

    /** In-cell bars, proportional to the value. */
    private static Workbook dataBars() {
        Workbook wb = figures();
        Sheet sheet = wb.sheet(0);

        sheet.range(1, 2, ROWS, 2).addConditionalFormat(ConditionalFormat.builder()
                .dataBar(Color.rgb(0x63, 0x8E, 0xC6))
                .gradient(true)
                .barBorderColor(Color.rgb(0x1F, 0x4E, 0x79))
                .showValue(true)
                .build());

        sheet.range(1, 4, ROWS, 4).addConditionalFormat(ConditionalFormat.builder()
                .dataBar(Color.rgb(0x63, 0xBE, 0x7B))
                .negativeFillColor(Color.rgb(0xC0, 0x00, 0x00))
                .axisColor(Color.BLACK)
                .axisPosition("midpoint")
                .showValue(true)
                .build());

        note(sheet, "Units get a plain bar; variance gets a midpoint axis so "
                + "negatives run left.");
        return wb;
    }

    /** A heat map across the revenue column. */
    private static Workbook colorScale() {
        Workbook wb = figures();
        Sheet sheet = wb.sheet(0);

        sheet.range(1, 3, ROWS, 3).addConditionalFormat(ConditionalFormat.builder()
                .colorScale(Color.rgb(0xF8, 0x69, 0x6B),
                        Color.rgb(0xFF, 0xEB, 0x84),
                        Color.rgb(0x63, 0xBE, 0x7B))
                .build());

        note(sheet, "A three-colour scale interpolates between the extremes of the range.");
        return wb;
    }

    /** Traffic lights, bucketed by the engine. */
    private static Workbook iconSet() {
        Workbook wb = figures();
        Sheet sheet = wb.sheet(0);

        sheet.range(1, 3, ROWS, 3).addConditionalFormat(ConditionalFormat.builder()
                .iconSet("3TrafficLights1")
                .build());

        note(sheet, "iconSet(\"3TrafficLights1\") — the thresholds default to thirds.");
        return wb;
    }

    /** A formula rule styles the whole row from one column's value. */
    private static Workbook rowRule() {
        Workbook wb = figures();
        Sheet sheet = wb.sheet(0);

        // $F2 is column-absolute, row-relative: the rule walks down the rows.
        sheet.range(1, 0, ROWS, 5).addConditionalFormat(ConditionalFormat.builder()
                .expression("$F2=\"overdue\"")
                .fillColor(Color.rgb(0xFF, 0xEB, 0x9C))
                .build());

        note(sheet, "expression(\"$F2=\\\"overdue\\\"\") — anchored column, relative row, "
                + "so the whole row lights up.");
        return wb;
    }

    private static void note(Sheet sheet, String text) {
        sheet.cell(ROWS + 2, 0).setValue(CellValue.text(text));
        sheet.range(ROWS + 2, 0, ROWS + 2, 0).setStyle(CellStyle.builder().italic().build());
    }

    /** The starting sheet: monthly figures read from a packaged workbook. */
    private static Workbook figures() {
        return Workbook.openBytes(DemoBooks.read(CONTENT), "xlsx");
    }
}
