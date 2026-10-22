package io.keikai.axyra.examples.featuredemo;

import io.keikai.axyra.examples.Example;
import io.keikai.axyra.sheets.CellValue;
import io.keikai.axyra.sheets.Range;
import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.Workbook;
import io.keikai.axyra.sheets.style.CellStyle;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.nio.file.Path;

/**
 * Open an existing workbook — from a stream, from bytes, and with a look at
 * what the reader found.
 *
 * <p>Opening is <b>lazy</b>: nothing is recalculated until you ask for it, so a
 * file with stale cached results keeps them until {@code recalculate()} runs.
 */
public final class OpenWorkbookExample implements Example {
    private static final String BOOK = "budget.xlsx";

    @Override
    public String name() {
        return "feature-demo/open-workbook";
    }

    @Override
    public String description() {
        return "Open a workbook from a stream or bytes, inspect it, and recalculate on demand";
    }

    @Override
    public void run(Path outputDirectory) throws IOException {
        try (Workbook wb = openFromStream()) {
            wb.save(outputDirectory.resolve("open-workbook-from-stream.xlsx"));
        }
        try (Workbook wb = openFromBytes()) {
            wb.save(outputDirectory.resolve("open-workbook-from-bytes.xlsx"));
        }
        try (Workbook wb = inspect()) {
            wb.save(outputDirectory.resolve("open-workbook-inspect.xlsx"));
        }
        try (Workbook wb = openAndRecalculate()) {
            wb.save(outputDirectory.resolve("open-workbook-recalculated.xlsx"));
        }
        System.out.println("Opened " + BOOK + " four ways; results in " + outputDirectory);
    }

    /** The reader is chosen from the format token, so any source works. */
    private static Workbook openFromStream() throws IOException {
        return Workbook.open(new ByteArrayInputStream(DemoBooks.read(BOOK)), "xlsx");
    }

    /** Straight from a byte[] — no temporary file, no stream to close. */
    private static Workbook openFromBytes() {
        return Workbook.openBytes(DemoBooks.read(BOOK), "xlsx");
    }

    /**
     * What the reader found: sheet names, indexes and used extents, written onto
     * a sheet of their own.
     */
    private static Workbook inspect() {
        Workbook wb = openFromBytes();
        Sheet summary = wb.createSheet("Opened");
        summary.range("A1:D1").setValues(new CellValue[][] {{
                CellValue.text("Index"), CellValue.text("Sheet"),
                CellValue.text("Used range"), CellValue.text("Cells")}});
        summary.range("A1:D1").setStyle(CellStyle.builder().bold().build());

        int row = 1;
        for (int i = 0; i < wb.sheetCount(); i++) {
            Sheet sheet = wb.sheet(i);
            if (sheet.index() == summary.index()) {
                continue;
            }
            Range used = sheet.usedRange();
            summary.setNumber(row, 0, sheet.index());
            summary.cell(row, 1).setValue(CellValue.text(sheet.name()));
            summary.cell(row, 2).setValue(CellValue.text(used == null ? "(empty)" : address(used)));
            summary.setNumber(row, 3, used == null ? 0 : (long) used.rowCount() * used.colCount());
            row++;
        }
        summary.setColumnWidth(1, 24);
        summary.setColumnWidth(2, 16);
        wb.setActiveSheet(summary.index());
        return wb;
    }

    /** Cached results survive the open; recalculation is opt-in. */
    private static Workbook openAndRecalculate() {
        Workbook wb = openFromBytes();
        wb.recalculate();
        return wb;
    }

    /** Formats a range as an A1 address such as {@code A1:D12}. */
    private static String address(Range r) {
        return col(r.firstColumn()) + (r.firstRow() + 1) + ":" + col(r.lastColumn()) + (r.lastRow() + 1);
    }

    private static String col(int index) {
        StringBuilder sb = new StringBuilder();
        for (int c = index + 1; c > 0; c = (c - 1) / 26) {
            sb.insert(0, (char) ('A' + (c - 1) % 26));
        }
        return sb.toString();
    }
}
