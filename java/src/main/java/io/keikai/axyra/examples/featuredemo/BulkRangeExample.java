package io.keikai.axyra.examples.featuredemo;

import io.keikai.axyra.examples.Example;
import io.keikai.axyra.sheets.CellValue;
import io.keikai.axyra.sheets.Range;
import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.Workbook;
import io.keikai.axyra.sheets.style.CellStyle;

import java.nio.file.Path;

/**
 * Bulk range read and write — the central design choice of the Java API.
 *
 * <p>{@code setValues} / {@code setNumbers} / {@code values()} / {@code numbers()}
 * cross the JNI boundary <b>once per call</b>, not once per cell. A 20x5 block is
 * one crossing, not a hundred; the {@code double[]} path skips boxing entirely.
 * Writing the same block cell by cell is the thing to avoid.
 */
public final class BulkRangeExample implements Example {
    private static final int ROWS = 20;

    @Override
    public String name() {
        return "feature-demo/bulk-range";
    }

    @Override
    public String description() {
        return "Write and read whole blocks in one call, including the unboxed numeric path";
    }

    @Override
    public void run(Path outputDirectory) {
        try (Workbook wb = mixedBlock()) {
            wb.save(outputDirectory.resolve("bulk-range-mixed-block.xlsx"));
        }
        try (Workbook wb = numericFastPath()) {
            wb.save(outputDirectory.resolve("bulk-range-numeric.xlsx"));
        }
        try (Workbook wb = readBlockBack()) {
            wb.save(outputDirectory.resolve("bulk-range-read-back.xlsx"));
        }
        System.out.println("Wrote three bulk-range workbooks to " + outputDirectory);
    }

    /** One {@code CellValue[][]}, one crossing, mixed types. */
    private static Workbook mixedBlock() {
        Workbook wb = Workbook.create();
        Sheet sheet = wb.createSheet("Block");

        CellValue[][] block = new CellValue[ROWS + 1][4];
        block[0] = new CellValue[] {CellValue.text("Order"), CellValue.text("Units"),
                CellValue.text("Priority"), CellValue.text("Cleared")};
        for (int i = 0; i < ROWS; i++) {
            block[i + 1] = new CellValue[] {
                    CellValue.text(String.format("ORD-%04d", 1001 + i)),
                    CellValue.number(20 + (i * 7) % 130),
                    CellValue.text(i % 3 == 0 ? "High" : i % 3 == 1 ? "Normal" : "Low"),
                    CellValue.bool(i % 4 != 0)};
        }
        sheet.range(0, 0, ROWS, 3).setValues(block);
        sheet.range("A1:D1").setStyle(CellStyle.builder().bold().build());
        sheet.setColumnWidth(0, 14);
        sheet.setColumnWidth(2, 12);
        sheet.setColumnWidth(3, 12);

        note(sheet, 0, 5, "One setValues call wrote " + (ROWS + 1) * 4 + " cells "
                + "in a single JNI crossing.");
        return wb;
    }

    /** A flat {@code double[]}, no boxing — the fastest way to fill numbers. */
    private static Workbook numericFastPath() {
        Workbook wb = Workbook.create();
        Sheet sheet = wb.createSheet("Numbers");

        sheet.range("A1:C1").setValues(new CellValue[][] {{
                CellValue.text("Month"), CellValue.text("Plan"), CellValue.text("Actual")}});
        sheet.range("A1:C1").setStyle(CellStyle.builder().bold().build());

        String[] months = {"Jan", "Feb", "Mar", "Apr", "May", "Jun",
                "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"};
        for (int i = 0; i < months.length; i++) {
            sheet.cell(i + 1, 0).setValue(CellValue.text(months[i]));
        }

        // setNumbers fills row-major across the whole rectangle.
        double[] figures = new double[months.length * 2];
        for (int i = 0; i < months.length; i++) {
            figures[i * 2] = 10_000 + i * 450;
            figures[i * 2 + 1] = 9_400 + ((i * 1373) % 3100);
        }
        sheet.range(1, 1, months.length, 2).setNumbers(figures);
        sheet.range(1, 1, months.length, 2).setStyle(
                CellStyle.builder().numberFormat("#,##0").build());

        note(sheet, 0, 4, "setNumbers(double[]) writes the rectangle row-major, unboxed.");
        return wb;
    }

    /** Reading is symmetrical: one call returns the whole rectangle. */
    private static Workbook readBlockBack() {
        Workbook wb = numericFastPath();
        Sheet sheet = wb.sheet(0);

        Range block = sheet.range(1, 1, 12, 2);
        double[] read = block.numbers();

        double planTotal = 0;
        double actualTotal = 0;
        for (int i = 0; i < read.length; i += 2) {
            planTotal += read[i];
            actualTotal += read[i + 1];
        }

        sheet.cell(14, 0).setValue(CellValue.text("Read back"));
        sheet.range("A15").setStyle(CellStyle.builder().bold().build());
        sheet.cell(15, 0).setValue(CellValue.text("values read"));
        sheet.setNumber(15, 1, read.length);
        sheet.cell(16, 0).setValue(CellValue.text("plan total"));
        sheet.setNumber(16, 1, planTotal);
        sheet.cell(17, 0).setValue(CellValue.text("actual total"));
        sheet.setNumber(17, 1, actualTotal);
        sheet.cell(18, 0).setValue(CellValue.text("variance"));
        sheet.setNumber(18, 1, actualTotal - planTotal);
        sheet.range(15, 1, 18, 1).setStyle(CellStyle.builder().numberFormat("#,##0").build());

        note(sheet, 20, 0, "Totals computed in Java from one numbers() call — "
                + "no per-cell reads, no formulas involved.");
        return wb;
    }

    private static void note(Sheet sheet, int row, int col, String text) {
        sheet.cell(row, col).setValue(CellValue.text(text));
        sheet.range(row, col, row, col).setStyle(CellStyle.builder().italic().build());
    }
}
