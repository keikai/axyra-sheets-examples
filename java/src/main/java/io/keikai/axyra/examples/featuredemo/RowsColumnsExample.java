package io.keikai.axyra.examples.featuredemo;

import io.keikai.axyra.examples.Example;
import io.keikai.axyra.sheets.CellValue;
import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.Workbook;
import io.keikai.axyra.sheets.style.CellStyle;
import io.keikai.axyra.sheets.style.Color;

import java.nio.file.Path;

/**
 * Rows and columns — sizing, hiding, autofit, insert and delete.
 *
 * <p>Row heights are in <b>points</b>, column widths in <b>character widths</b>,
 * which is Excel's own pair of units. Inserting or deleting shifts the cells
 * below or to the right and rewrites the formulas that referenced them.
 */
public final class RowsColumnsExample implements Example {
    private static final String CONTENT = "rowsColumns-data-content.xlsx";

    @Override
    public String name() {
        return "feature-demo/rows-columns";
    }

    @Override
    public String description() {
        return "Size, autofit, hide, insert and delete rows and columns";
    }

    @Override
    public void run(Path outputDirectory) {
        try (Workbook wb = sizes()) {
            wb.save(outputDirectory.resolve("rows-columns-sizes.xlsx"));
        }
        try (Workbook wb = autofit()) {
            wb.save(outputDirectory.resolve("rows-columns-autofit.xlsx"));
        }
        try (Workbook wb = hide()) {
            wb.save(outputDirectory.resolve("rows-columns-hidden.xlsx"));
        }
        try (Workbook wb = insertDelete()) {
            wb.save(outputDirectory.resolve("rows-columns-insert.xlsx"));
        }
        System.out.println("Wrote four row and column workbooks to " + outputDirectory);
    }

    /** Explicit sizes, and the sheet-wide defaults everything else inherits. */
    private static Workbook sizes() {
        Workbook wb = data("Sizes");
        Sheet sheet = wb.sheet(0);

        sheet.setRowHeight(0, 34);
        sheet.setColumnWidth(0, 26);
        sheet.setColumnWidth(1, 6);
        sheet.setDefaultRowHeight(18);

        int row = 8;
        sheet.cell(row, 0).setValue(CellValue.text("Read back"));
        sheet.range(row, 0, row, 0).setStyle(CellStyle.builder().bold().build());
        report(sheet, ++row, "rowHeight(0)", sheet.rowHeight(0) + " pt");
        report(sheet, ++row, "isRowHeightCustom(0)", String.valueOf(sheet.isRowHeightCustom(0)));
        report(sheet, ++row, "columnWidth(0)", sheet.columnWidth(0) + " chars");
        report(sheet, ++row, "isColumnWidthCustom(2)", String.valueOf(sheet.isColumnWidthCustom(2)));
        report(sheet, ++row, "defaultRowHeight()", String.valueOf(sheet.defaultRowHeight()));
        return wb;
    }

    /** Autofit measures the rendered text, so it needs the content first. */
    private static Workbook autofit() {
        Workbook wb = Workbook.create();
        Sheet sheet = wb.createSheet("Autofit");

        sheet.cell(0, 0).setValue(CellValue.text("Column A was autofitted"));
        sheet.cell(1, 0).setValue(CellValue.text("A considerably longer line of text"));
        sheet.cell(2, 0).setValue(CellValue.text("short"));
        sheet.cell(0, 2).setValue(CellValue.text("Column C was left at the default"));
        sheet.cell(1, 2).setValue(CellValue.text("A considerably longer line of text"));

        sheet.autoFitColumn(0);
        sheet.autoFitRow(1);

        // Reported in G1, clear of the columns being measured: in column A it would
        // itself be the longest string there.
        sheet.cell(0, 6).setValue(CellValue.text(
                "autoFitColumn(0) -> " + sheet.columnWidth(0) + " character widths"));
        return wb;
    }

    /** Hiding keeps the data; it only sets the row or column flag. */
    private static Workbook hide() {
        Workbook wb = data("Hidden");
        Sheet sheet = wb.sheet(0);

        sheet.setRowHidden(3, true);
        sheet.setColumnHidden(2, true);

        sheet.cell(8, 0).setValue(CellValue.text("Row 4 and column C are hidden."));
        sheet.cell(9, 0).setValue(CellValue.text(
                "isRowHidden(3) = " + sheet.isRowHidden(3)
                        + ", isColumnHidden(2) = " + sheet.isColumnHidden(2)));
        sheet.cell(10, 0).setValue(CellValue.text(
                "The SUM below still counts the hidden row — hiding is presentation only."));
        sheet.cell(11, 0).setValue(CellValue.text("SUM(B2:B6)"));
        sheet.cell(11, 1).setFormula("SUM(B2:B6)");
        sheet.setColumnWidth(0, 34);
        wb.recalculate();
        return wb;
    }

    /** Insert and delete shift the neighbours and repair the formulas. */
    private static Workbook insertDelete() {
        Workbook wb = data("Shifted");
        Sheet sheet = wb.sheet(0);
        sheet.cell(7, 0).setValue(CellValue.text("Total"));
        sheet.cell(7, 1).setFormula("SUM(B2:B6)");
        sheet.range("A8:B8").setStyle(CellStyle.builder().bold().build());

        sheet.insertRows(2, 2);
        sheet.cell(2, 0).setValue(CellValue.text("<- two rows inserted here"));
        sheet.range("A3").setStyle(CellStyle.builder()
                .italic().fontColor(Color.rgb(0x1F, 0x4E, 0x79)).build());

        sheet.insertColumns(1, 1);
        sheet.cell(0, 1).setValue(CellValue.text("Inserted"));

        sheet.cell(12, 0).setValue(CellValue.text("The total's formula is now"));
        sheet.cell(13, 0).setValue(CellValue.text(sheet.cell(9, 2).formula()));
        sheet.setColumnWidth(0, 30);
        wb.recalculate();
        return wb;
    }

    private static void report(Sheet sheet, int row, String call, String result) {
        sheet.cell(row, 0).setValue(CellValue.text(call));
        sheet.cell(row, 1).setValue(CellValue.text(result));
    }

    /** The starting sheet: a small data table read from a packaged workbook, renamed for the sample. */
    private static Workbook data(String name) {
        Workbook wb = Workbook.openBytes(DemoBooks.read(CONTENT), "xlsx");
        wb.renameSheet(0, name);
        return wb;
    }
}
