package io.keikai.axyra.examples.featuredemo;

import io.keikai.axyra.examples.Example;
import io.keikai.axyra.sheets.CellValue;
import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.Workbook;
import io.keikai.axyra.sheets.content.Chart;
import io.keikai.axyra.sheets.content.LegendPosition;
import io.keikai.axyra.sheets.style.Color;

import java.nio.file.Path;

/**
 * Insert a chart.
 *
 * <p>A chart is declared in one builder expression: type, series, legend,
 * grouping and the anchor rectangle it occupies on the sheet. Series are bound
 * to ranges by A1 address, so the chart redraws whenever those cells change.
 */
public final class InsertChartExample implements Example {
    private static final String CONTENT = "insertChart-data-content.xlsx";

    @Override
    public String name() {
        return "feature-demo/insert-chart";
    }

    @Override
    public String description() {
        return "Insert column, line, area and pie charts bound to sheet ranges";
    }

    @Override
    public void run(Path outputDirectory) {
        try (Workbook wb = column()) {
            wb.save(outputDirectory.resolve("insert-chart-column.xlsx"));
        }
        try (Workbook wb = line()) {
            wb.save(outputDirectory.resolve("insert-chart-line.xlsx"));
        }
        try (Workbook wb = area()) {
            wb.save(outputDirectory.resolve("insert-chart-area.xlsx"));
        }
        try (Workbook wb = pie()) {
            wb.save(outputDirectory.resolve("insert-chart-pie.xlsx"));
        }
        try (Workbook wb = twoCharts()) {
            wb.save(outputDirectory.resolve("insert-chart-two-charts.xlsx"));
        }
        System.out.println("Wrote five chart workbooks to " + outputDirectory);
    }

    private static Workbook column() {
        Workbook wb = data();
        Sheet sheet = wb.sheet(0);
        sheet.addChart(Chart.column()
                .title("Plan vs actual")
                .series("Plan", "Data!B2:B7", "Data!A2:A7")
                .series("Actual", "Data!C2:C7", "Data!A2:A7", Color.rgb(0x1F, 0x4E, 0x79))
                .legend(LegendPosition.BOTTOM)
                .grouping("clustered")
                .gapWidth(60)
                .anchor(1, 4, 18, 12)
                .build());
        return wb;
    }

    private static Workbook line() {
        Workbook wb = data();
        Sheet sheet = wb.sheet(0);
        sheet.addChart(Chart.line()
                .title("Trend")
                .series("Plan", "Data!B2:B7", "Data!A2:A7")
                .series("Actual", "Data!C2:C7", "Data!A2:A7")
                .legend(LegendPosition.RIGHT)
                .anchor(1, 4, 18, 12)
                .build());
        return wb;
    }

    private static Workbook area() {
        Workbook wb = data();
        Sheet sheet = wb.sheet(0);
        sheet.addChart(Chart.area()
                .title("Cumulative")
                .series("Plan", "Data!B2:B7", "Data!A2:A7")
                .series("Actual", "Data!C2:C7", "Data!A2:A7")
                .grouping("stacked")
                .legend(LegendPosition.BOTTOM)
                .anchor(1, 4, 18, 12)
                .build());
        return wb;
    }

    private static Workbook pie() {
        Workbook wb = data();
        Sheet sheet = wb.sheet(0);
        sheet.addChart(Chart.pie()
                .title("Actual by month")
                .series("Actual", "Data!C2:C7", "Data!A2:A7")
                .legend(LegendPosition.RIGHT)
                .anchor(1, 4, 18, 12)
                .build());
        return wb;
    }

    /** Charts are anchored to a cell rectangle, so several fit on one sheet. */
    private static Workbook twoCharts() {
        Workbook wb = data();
        Sheet sheet = wb.sheet(0);

        sheet.addChart(Chart.column()
                .title("Plan vs actual")
                .series("Plan", "Data!B2:B7", "Data!A2:A7")
                .series("Actual", "Data!C2:C7", "Data!A2:A7")
                .legend(LegendPosition.BOTTOM)
                .anchor(1, 4, 15, 10)
                .build());

        sheet.addChart(Chart.pie()
                .title("Share of actual")
                .series("Actual", "Data!C2:C7", "Data!A2:A7")
                .legend(LegendPosition.RIGHT)
                .anchor(16, 4, 30, 10)
                .build());

        sheet.cell(9, 0).setValue(CellValue.text(
                "anchor(firstRow, firstCol, lastRow, lastCol) places each one."));
        return wb;
    }

    /** The starting sheet: a "Data" sheet of monthly plan and actual figures. */
    private static Workbook data() {
        return Workbook.openBytes(DemoBooks.read(CONTENT), "xlsx");
    }
}
