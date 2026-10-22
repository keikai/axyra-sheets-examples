package io.keikai.axyra.examples.featuredemo;

import io.keikai.axyra.examples.Example;
import io.keikai.axyra.sheets.CellValue;
import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.Workbook;
import io.keikai.axyra.sheets.content.Validation;
import io.keikai.axyra.sheets.style.BorderStyle;
import io.keikai.axyra.sheets.style.CellStyle;
import io.keikai.axyra.sheets.style.Color;

import java.nio.file.Path;
import java.time.LocalDate;

/**
 * Data validation — constrain what a cell will accept.
 *
 * <p>Each sample writes a rule over a range and shades that range, so the
 * governed area is legible while it is still empty. {@link Validation} covers
 * list, whole number, decimal, date, time, text length and arbitrary formula
 * rules. A list rule shows its in-cell arrow unless {@code showDropdown(false)}
 * asks otherwise.
 */
public final class DataValidationExample implements Example {
    @Override
    public String name() {
        return "feature-demo/data-validation";
    }

    @Override
    public String description() {
        return "Add dropdown, numeric, date, prompt and custom-formula validation rules";
    }

    @Override
    public void run(Path outputDirectory) {
        try (Workbook wb = dropdown()) {
            wb.save(outputDirectory.resolve("data-validation-dropdown.xlsx"));
        }
        try (Workbook wb = numeric()) {
            wb.save(outputDirectory.resolve("data-validation-numeric.xlsx"));
        }
        try (Workbook wb = prompts()) {
            wb.save(outputDirectory.resolve("data-validation-prompts.xlsx"));
        }
        try (Workbook wb = custom()) {
            wb.save(outputDirectory.resolve("data-validation-custom.xlsx"));
        }
        System.out.println("Wrote four data-validation workbooks to " + outputDirectory);
    }

    /** An inline list, and one backed by a range of allowed values. */
    private static Workbook dropdown() {
        Workbook wb = Workbook.create();
        Sheet sheet = form(wb, "Region", "Stage");

        sheet.range("A2:A20").setValidation(
                Validation.list("North", "South", "East", "West"));

        // The allowed values live on the sheet, so they are editable data.
        sheet.cell(0, 4).setValue(CellValue.text("Allowed stages"));
        sheet.range("E2:E5").setValues(new CellValue[][] {
                {CellValue.text("Draft")}, {CellValue.text("Review")},
                {CellValue.text("Final")}, {CellValue.text("Void")}});
        sheet.range("B2:B20").setValidation(Validation.builder()
                .type("list")
                .formula1("$E$2:$E$5")
                .build());

        markInputs(sheet, "A2:A20", "B2:B20");
        hint(sheet, "Columns A and B each carry a dropdown; A lists its values inline, "
                + "B reads them from E2:E5.");
        return wb;
    }

    /** Ranges, comparisons and dates, all as format-free criteria strings. */
    private static Workbook numeric() {
        Workbook wb = Workbook.create();
        Sheet sheet = form(wb, "Quantity", "Delivery date");

        sheet.range("A2:A20").setValidation(
                Validation.wholeNumber("between", "1", "100"));
        sheet.range("B2:B20").setValidation(
                Validation.date("between", "DATE(2026,1,1)", "DATE(2026,12,31)"));
        sheet.range("C2:C20").setValidation(
                Validation.decimal("greaterThanOrEqual", "0", null));
        sheet.range("D2:D20").setValidation(
                Validation.textLength("lessThanOrEqual", "40", null));

        sheet.cell(0, 2).setValue(CellValue.text("Unit price"));
        sheet.cell(0, 3).setValue(CellValue.text("Comment"));
        sheet.range("A1:D1").setStyle(CellStyle.builder().bold().build());

        // One row that satisfies every rule, as an example of what each column accepts.
        sheet.setNumber(1, 0, 25);
        sheet.cell(1, 1).setDate(LocalDate.of(2026, 6, 15));
        sheet.setNumber(1, 2, 19.99);
        sheet.cell(1, 3).setValue(CellValue.text("Ship before month end"));

        markInputs(sheet, "A2:A20", "B2:B20", "C2:C20", "D2:D20");
        sheet.range("B2:B20").setStyle(inputStyle().numberFormat("yyyy-mm-dd").build());
        sheet.range("C2:C20").setStyle(inputStyle().numberFormat("#,##0.00").build());

        hint(sheet, "Quantity 1-100, delivery inside 2026, price >= 0, comment <= 40 chars. "
                + "Row 2 holds a value each rule accepts.");
        return wb;
    }

    /** The message a user sees before typing, and the one that stops them. */
    private static Workbook prompts() {
        Workbook wb = Workbook.create();
        Sheet sheet = form(wb, "Stage", "Headcount");

        sheet.range("B2:B5").setValues(new CellValue[][] {
                {CellValue.text("Draft")}, {CellValue.text("Review")},
                {CellValue.text("Final")}, {CellValue.text("Void")}});
        sheet.setColumnHidden(1, true);

        sheet.range("A2:A20").setValidation(Validation.builder()
                .type("list")
                .formula1("$B$2:$B$5")
                .allowBlank(false)
                .inputMessage("Workflow stage", "Pick a stage from the list.")
                .errorMessage("Not a stage", "Use the dropdown — free text is rejected.")
                .errorStyle("stop")
                .build());

        sheet.cell(1, 0).setValue(CellValue.text("Draft"));
        markInputs(sheet, "A2:A20");

        hint(sheet, "\"stop\" blocks the entry. \"warning\" and \"information\" let it through.");
        return wb;
    }

    /** Anything expressible as a formula can be a rule. */
    private static Workbook custom() {
        Workbook wb = Workbook.create();
        Sheet sheet = form(wb, "Multiple of 5", "Order ref");

        sheet.range("A2:A20").setValidation(
                Validation.custom("AND(ISNUMBER(A2), MOD(A2,5)=0)"));
        sheet.range("B2:B20").setValidation(
                Validation.custom("LEFT(B2,4)=\"ORD-\""));

        sheet.setNumber(1, 0, 25);
        sheet.cell(1, 1).setValue(CellValue.text("ORD-1234"));
        markInputs(sheet, "A2:A20", "B2:B20");

        hint(sheet, "Column A accepts multiples of five; column B needs an ORD- prefix. "
                + "The formula is relative to the range's first cell.");
        return wb;
    }

    private static Sheet form(Workbook wb, String first, String second) {
        Sheet sheet = wb.createSheet("Entry");
        sheet.cell(0, 0).setValue(CellValue.text(first));
        sheet.cell(0, 1).setValue(CellValue.text(second));
        sheet.range("A1:B1").setStyle(CellStyle.builder()
                .bold().fill(Color.rgb(0xDD, 0xEB, 0xF7)).build());
        sheet.setColumnWidth(0, 20);
        sheet.setColumnWidth(1, 20);
        sheet.setColumnWidth(2, 16);
        sheet.setColumnWidth(3, 20);
        sheet.setColumnWidth(4, 18);
        return sheet;
    }

    /**
     * The input-cell marking: a {@code #FFF2CC} fill inside a {@code #D9A441} edge.
     * Returned as a builder so a column can add its own number format without
     * losing the marking.
     */
    private static CellStyle.Builder inputStyle() {
        return CellStyle.builder()
                .fill(Color.rgb(0xFF, 0xF2, 0xCC))
                .allBorders(BorderStyle.THIN, Color.rgb(0xD9, 0xA4, 0x41));
    }

    /**
     * Marks the ranges a rule governs. A validation is invisible until it rejects
     * something, so shading the governed cells makes them legible while empty.
     */
    private static void markInputs(Sheet sheet, String... ranges) {
        for (String ref : ranges) {
            sheet.range(ref).setStyle(inputStyle().build());
        }
    }

    private static void hint(Sheet sheet, String text) {
        sheet.cell(22, 0).setValue(CellValue.text(text));
        sheet.range(22, 0, 22, 0).setStyle(CellStyle.builder().italic().build());
    }
}
