package io.keikai.axyra.examples.featuredemo;

import io.keikai.axyra.examples.Example;
import io.keikai.axyra.sheets.CellValue;
import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.Workbook;
import io.keikai.axyra.sheets.style.CellStyle;

import java.nio.file.Path;

/**
 * Create a table — Excel's "Format as Table", from Java.
 *
 * <p>A table is more than banded styling: its columns get names, and those names
 * become <b>structured references</b>, so {@code SUM(Orders[Amount])} keeps
 * working when rows are added. Header and totals rows are toggled by name.
 */
public final class CreateTableExample implements Example {
    private static final String CONTENT = "createTable-orders-content.xlsx";

    @Override
    public String name() {
        return "feature-demo/create-table";
    }

    @Override
    public String description() {
        return "Create a table, use structured references and toggle header and totals rows";
    }

    @Override
    public void run(Path outputDirectory) {
        try (Workbook wb = create()) {
            wb.save(outputDirectory.resolve("create-table-named.xlsx"));
        }
        try (Workbook wb = inferred()) {
            wb.save(outputDirectory.resolve("create-table-inferred.xlsx"));
        }
        try (Workbook wb = structured()) {
            wb.save(outputDirectory.resolve("create-table-structured.xlsx"));
        }
        try (Workbook wb = headerAndTotals()) {
            wb.save(outputDirectory.resolve("create-table-totals.xlsx"));
        }
        System.out.println("Wrote four table workbooks to " + outputDirectory);
    }

    /** Explicit column names, given at creation. */
    private static Workbook create() {
        Workbook wb = orders();
        Sheet sheet = wb.sheet(0);
        sheet.range(0, 0, 6, 3).createTable("Orders", "Item", "Qty", "Price", "Amount");
        note(sheet, 8, "createTable(\"Orders\", \"Item\", \"Qty\", \"Price\", \"Amount\")");
        wb.recalculate();
        return wb;
    }

    /** With no names given, the header row supplies them. */
    private static Workbook inferred() {
        Workbook wb = orders();
        Sheet sheet = wb.sheet(0);
        sheet.range(0, 0, 6, 3).createTable("Orders");

        // The names came from row 1, so structured references work immediately.
        note(sheet, 8, "createTable(\"Orders\") took the column names from row 1:");
        sheet.cell(9, 0).setValue(CellValue.text("SUM(Orders[Amount])"));
        sheet.cell(9, 3).setFormula("SUM(Orders[Amount])");
        sheet.range(9, 3, 9, 3).setStyle(CellStyle.builder()
                .bold().numberFormat("#,##0.00").build());
        wb.recalculate();
        return wb;
    }

    /** The point of naming the columns: formulas that survive edits. */
    private static Workbook structured() {
        Workbook wb = orders();
        Sheet sheet = wb.sheet(0);
        sheet.range(0, 0, 6, 3).createTable("Orders", "Item", "Qty", "Price", "Amount");

        sheet.cell(8, 0).setValue(CellValue.text("SUM(Orders[Amount])"));
        sheet.cell(8, 3).setFormula("SUM(Orders[Amount])");
        sheet.cell(9, 0).setValue(CellValue.text("AVERAGE(Orders[Price])"));
        sheet.cell(9, 3).setFormula("AVERAGE(Orders[Price])");
        sheet.cell(10, 0).setValue(CellValue.text("COUNTA(Orders[Item])"));
        sheet.cell(10, 3).setFormula("COUNTA(Orders[Item])");
        sheet.range(8, 0, 10, 0).setStyle(CellStyle.builder().bold().build());
        sheet.range(8, 3, 9, 3).setStyle(CellStyle.builder().numberFormat("#,##0.00").build());

        note(sheet, 12, "These reference the table by name, not by A1 address — "
                + "add a row and they still cover it.");
        wb.recalculate();
        return wb;
    }

    /** Both rows are sheet-level toggles, addressed by the table's name. */
    private static Workbook headerAndTotals() {
        Workbook wb = orders();
        Sheet sheet = wb.sheet(0);
        sheet.range(0, 0, 6, 3).createTable("Orders", "Item", "Qty", "Price", "Amount");

        sheet.setTableHeaderRow("Orders", true);
        sheet.setTableTotalsRow("Orders", true);

        note(sheet, 9, "setTableTotalsRow(\"Orders\", true) added the totals row below the data.");
        wb.recalculate();
        return wb;
    }

    private static void note(Sheet sheet, int row, String text) {
        sheet.cell(row, 0).setValue(CellValue.text(text));
        sheet.range(row, 0, row, 0).setStyle(CellStyle.builder().italic().build());
    }

    /** The starting sheet: order lines read from a packaged workbook. */
    private static Workbook orders() {
        return Workbook.openBytes(DemoBooks.read(CONTENT), "xlsx");
    }
}
