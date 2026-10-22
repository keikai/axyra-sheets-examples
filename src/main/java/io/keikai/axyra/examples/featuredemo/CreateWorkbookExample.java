package io.keikai.axyra.examples.featuredemo;

import io.keikai.axyra.examples.Example;
import io.keikai.axyra.sheets.CellValue;
import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.Workbook;
import io.keikai.axyra.sheets.style.CellStyle;

import java.nio.file.Path;

/**
 * Create a new workbook.
 *
 * <p>{@code Workbook.create()} builds an empty book on the sparse backend, the
 * general-purpose default. {@code createColumnar()} builds a behaviourally
 * identical one on columnar storage, which is far leaner in memory for large,
 * dense, numeric sheets. A book read from a file always uses the reader's
 * representation, so the choice only applies to books you generate.
 */
public final class CreateWorkbookExample implements Example {
    @Override
    public String name() {
        return "feature-demo/create-workbook";
    }

    @Override
    public String description() {
        return "Create workbooks on the sparse and columnar backends, and set workbook defaults";
    }

    @Override
    public void run(Path outputDirectory) {
        try (Workbook wb = sparse()) {
            wb.save(outputDirectory.resolve("create-workbook-sparse.xlsx"));
        }
        try (Workbook wb = columnar()) {
            wb.save(outputDirectory.resolve("create-workbook-columnar.xlsx"));
        }
        try (Workbook wb = defaults()) {
            wb.save(outputDirectory.resolve("create-workbook-defaults.xlsx"));
        }
        System.out.println("Created three workbooks in " + outputDirectory);
    }

    /** The general-purpose backend: a book from nothing, with one cell in it. */
    private static Workbook sparse() {
        Workbook wb = Workbook.create();
        Sheet sheet = wb.createSheet("Sheet1");
        sheet.cell(0, 0).setValue(CellValue.text("Hello"));

        sheet.cell(2, 0).setValue(CellValue.text(
                "Workbook.create() — sparse storage, the general-purpose default."));
        sheet.setColumnWidth(0, 60);
        return wb;
    }

    /**
     * The columnar backend trades random-access cost for a much smaller footprint
     * on wide, dense, numeric sheets.
     */
    private static Workbook columnar() {
        Workbook wb = Workbook.createColumnar();
        Sheet sheet = wb.createSheet("Wide");
        sheet.range("A1:Z1").setNumbers(new double[26]);

        sheet.cell(2, 0).setValue(CellValue.text(
                "Workbook.createColumnar() — 26 columns written in one call."));
        sheet.cell(3, 0).setValue(CellValue.text(
                "Same API and same behaviour as create(); leaner at rest on large, "
                        + "dense, numeric sheets."));
        sheet.setColumnWidth(0, 60);
        return wb;
    }

    /** Workbook-wide defaults, written onto the sheet rather than to the console. */
    private static Workbook defaults() {
        Workbook wb = Workbook.create();
        Sheet sheet = wb.createSheet("Defaults");

        wb.setDefaultFont(CellStyle.builder().fontName("Calibri").fontSize(11).build());

        sheet.range("A1:B1").setValues(new CellValue[][] {{
                CellValue.text("Property"), CellValue.text("Value")}});
        sheet.range("A1:B1").setStyle(CellStyle.builder().bold().build());

        sheet.cell(1, 0).setValue(CellValue.text("default font"));
        sheet.cell(1, 1).setValue(CellValue.text(
                wb.defaultStyle().fontName() + " " + wb.defaultStyle().fontSize()));
        sheet.cell(2, 0).setValue(CellValue.text("locale"));
        sheet.cell(2, 1).setValue(CellValue.text(wb.locale()));
        sheet.cell(3, 0).setValue(CellValue.text("1904 date system"));
        sheet.cell(3, 1).setValue(CellValue.bool(wb.date1904()));
        sheet.cell(4, 0).setValue(CellValue.text("sheets"));
        sheet.setNumber(4, 1, wb.sheetCount());

        sheet.setColumnWidth(0, 24);
        sheet.setColumnWidth(1, 24);
        return wb;
    }
}
