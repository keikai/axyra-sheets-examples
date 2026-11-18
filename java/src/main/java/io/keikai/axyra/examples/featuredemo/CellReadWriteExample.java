package io.keikai.axyra.examples.featuredemo;

import io.keikai.axyra.examples.Example;
import io.keikai.axyra.sheets.CellValue;
import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.Workbook;
import io.keikai.axyra.sheets.style.CellStyle;

import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalDateTime;

/**
 * Write and read a single cell.
 *
 * <p>{@link CellValue} is a sealed type — {@code Number}, {@code Text},
 * {@code Bool}, {@code Blank} and friends — so the kind that went in is the kind
 * that comes back. Dates go in as {@code java.time} values over the Excel serial
 * and come back the same way.
 */
public final class CellReadWriteExample implements Example {
    private static final String CONTENT = "cellReadWrite-header-content.xlsx";

    @Override
    public String name() {
        return "feature-demo/cell-read-write";
    }

    @Override
    public String description() {
        return "Write every CellValue kind, read it back, and round-trip dates and times";
    }

    @Override
    public void run(Path outputDirectory) {
        try (Workbook wb = writeTypes()) {
            wb.save(outputDirectory.resolve("cell-read-write-types.xlsx"));
        }
        try (Workbook wb = readBack()) {
            wb.save(outputDirectory.resolve("cell-read-write-read-back.xlsx"));
        }
        try (Workbook wb = datesAndTimes()) {
            wb.save(outputDirectory.resolve("cell-read-write-dates.xlsx"));
        }
        System.out.println("Wrote three cell read/write workbooks to " + outputDirectory);
    }

    /** Every {@code CellValue} kind, one per row. */
    private static Workbook writeTypes() {
        Workbook wb = header("Written");
        Sheet sheet = wb.sheet(0);

        sheet.cell(1, 0).setValue(CellValue.text("text"));
        sheet.cell(1, 1).setValue(CellValue.text("Widget"));

        sheet.cell(2, 0).setValue(CellValue.text("number"));
        sheet.cell(2, 1).setValue(CellValue.number(1234.5));

        sheet.cell(3, 0).setValue(CellValue.text("boolean"));
        sheet.cell(3, 1).setValue(CellValue.bool(true));

        sheet.cell(4, 0).setValue(CellValue.text("blank"));
        sheet.cell(4, 1).setValue(CellValue.blank());

        sheet.cell(5, 0).setValue(CellValue.text("formula"));
        sheet.cell(5, 1).setFormula("B3*2");

        // The shorthand for the common case: a bare number, no wrapper.
        sheet.cell(6, 0).setValue(CellValue.text("setNumber"));
        sheet.setNumber(6, 1, 42);

        wb.recalculate();
        return wb;
    }

    /** What comes back out, which variant carried it, and how it displays. */
    private static Workbook readBack() {
        Workbook wb = writeTypes();
        Sheet sheet = wb.sheet(0);
        sheet.range("C1:D1").setValues(new CellValue[][] {{
                CellValue.text("CellValue variant"), CellValue.text("Displayed as")}});
        sheet.range("C1:D1").setStyle(CellStyle.builder().bold().build());

        for (int row = 1; row <= 6; row++) {
            CellValue value = sheet.cell(row, 1).value();
            sheet.cell(row, 2).setValue(CellValue.text(value.getClass().getSimpleName()));
            sheet.cell(row, 3).setValue(CellValue.text(sheet.cell(row, 1).formattedText()));
        }
        sheet.setColumnWidth(2, 22);
        sheet.setColumnWidth(3, 22);
        sheet.cell(8, 0).setValue(CellValue.text(
                "Every value goes in and comes back out as the same kind."));
        return wb;
    }

    /** {@code setDate} / {@code setDateTime} write the Excel serial; the format makes it legible. */
    private static Workbook datesAndTimes() {
        Workbook wb = header("Dates");
        Sheet sheet = wb.sheet(0);

        sheet.cell(1, 0).setValue(CellValue.text("LocalDate"));
        sheet.cell(1, 1).setDate(LocalDate.of(2026, 8, 24));
        sheet.range("B2").setStyle(CellStyle.builder().numberFormat("yyyy-mm-dd").build());

        sheet.cell(2, 0).setValue(CellValue.text("LocalDateTime"));
        sheet.cell(2, 1).setDateTime(LocalDateTime.of(2026, 8, 24, 14, 30));
        sheet.range("B3").setStyle(CellStyle.builder().numberFormat("yyyy-mm-dd hh:mm").build());

        sheet.cell(3, 0).setValue(CellValue.text("The raw serial"));
        sheet.cell(3, 1).setValue(sheet.cell(1, 1).value());

        sheet.cell(5, 0).setValue(CellValue.text("dateValue() ->"));
        sheet.cell(5, 1).setValue(CellValue.text(String.valueOf(sheet.cell(1, 1).dateValue())));
        sheet.cell(6, 0).setValue(CellValue.text("dateTimeValue() ->"));
        sheet.cell(6, 1).setValue(CellValue.text(String.valueOf(sheet.cell(2, 1).dateTimeValue())));

        sheet.setColumnWidth(1, 22);
        return wb;
    }

    /** The starting sheet: a header row read from a packaged workbook, renamed for the sample. */
    private static Workbook header(String name) {
        Workbook wb = Workbook.openBytes(DemoBooks.read(CONTENT), "xlsx");
        wb.renameSheet(0, name);
        return wb;
    }
}
