package io.keikai.axyra.examples.featuredemo;

import io.keikai.axyra.examples.Example;
import io.keikai.axyra.sheets.CellValue;
import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.Workbook;
import io.keikai.axyra.sheets.format.Formats;
import io.keikai.axyra.sheets.format.ParsedValue;
import io.keikai.axyra.sheets.style.CellStyle;

import java.nio.file.Path;
import java.time.LocalDate;

/**
 * Number and date formats — Excel format codes, applied and evaluated.
 *
 * <p>A format never changes the stored value, only its presentation:
 * {@code formattedText()} returns what a viewer shows, {@code value()} the
 * number underneath. {@link Formats} does the same job statelessly, with no
 * workbook involved, and both are locale-aware.
 */
public final class NumberFormatsExample implements Example {
    @Override
    public String name() {
        return "feature-demo/number-formats";
    }

    @Override
    public String description() {
        return "Apply number and date format codes, per locale and without a workbook";
    }

    @Override
    public void run(Path outputDirectory) {
        try (Workbook wb = numbers()) {
            wb.save(outputDirectory.resolve("number-formats-numbers.xlsx"));
        }
        try (Workbook wb = dates()) {
            wb.save(outputDirectory.resolve("number-formats-dates.xlsx"));
        }
        try (Workbook wb = locales()) {
            wb.save(outputDirectory.resolve("number-formats-locales.xlsx"));
        }
        try (Workbook wb = stateless()) {
            wb.save(outputDirectory.resolve("number-formats-stateless.xlsx"));
        }
        System.out.println("Wrote four number-format workbooks to " + outputDirectory);
    }

    private static Workbook numbers() {
        Workbook wb = Workbook.create();
        Sheet sheet = header(wb, "Numbers", "Format code", "Formatted", "Underlying value");

        String[] codes = {
                "General", "#,##0", "#,##0.00", "0.0%", "0.0%;[Red]-0.0%",
                "$#,##0.00", "[$€-407]#,##0.00", "0.00E+00", "# ?/?", "\"OK\";\"OK\";\"zero\""};
        double[] values = {1234.5678, 1234.5678, 1234.5678, 0.0742, -0.0742,
                1234.5678, 1234.5678, 1234.5678, 1.75, 1};

        for (int i = 0; i < codes.length; i++) {
            int row = i + 1;
            sheet.cell(row, 0).setValue(CellValue.text(codes[i]));
            sheet.setNumber(row, 1, values[i]);
            sheet.range(row, 1, row, 1).setStyle(
                    CellStyle.builder().numberFormat(codes[i]).build());
            sheet.cell(row, 2).setValue(CellValue.text(String.valueOf(values[i])));
        }
        return wb;
    }

    private static Workbook dates() {
        Workbook wb = Workbook.create();
        Sheet sheet = header(wb, "Dates", "Format code", "Formatted", "Serial");

        String[] codes = {
                "yyyy-mm-dd", "d mmm yyyy", "dddd d mmmm yyyy",
                "mm/dd/yy", "mmm-yy", "hh:mm:ss", "yyyy-mm-dd hh:mm"};
        LocalDate date = LocalDate.of(2026, 8, 24);

        for (int i = 0; i < codes.length; i++) {
            int row = i + 1;
            sheet.cell(row, 0).setValue(CellValue.text(codes[i]));
            sheet.cell(row, 1).setDate(date);
            sheet.range(row, 1, row, 1).setStyle(
                    CellStyle.builder().numberFormat(codes[i]).build());
            sheet.cell(row, 2).setValue(sheet.cell(row, 1).value());
        }
        sheet.cell(codes.length + 2, 0).setValue(CellValue.text(
                "Every row holds the same serial — only the format differs."));
        return wb;
    }

    /** Month names, separators and grouping follow the locale. */
    private static Workbook locales() {
        Workbook wb = Workbook.create();
        Sheet sheet = header(wb, "Locales", "Locale", "dddd d mmmm yyyy", "#,##0.00");

        String[] tags = {"en-US", "de-DE", "fr-FR", "ja-JP"};
        for (int i = 0; i < tags.length; i++) {
            int row = i + 1;
            sheet.cell(row, 0).setValue(CellValue.text(tags[i]));
            sheet.cell(row, 1).setValue(CellValue.text(
                    Formats.format(46_258, "dddd d mmmm yyyy", tags[i])));
            sheet.cell(row, 2).setValue(CellValue.text(
                    Formats.format(1234.5678, "#,##0.00", tags[i])));
        }
        sheet.setColumnWidth(1, 30);
        return wb;
    }

    /** {@link Formats} formats and parses without a workbook at all. */
    private static Workbook stateless() {
        Workbook wb = Workbook.create();
        Sheet sheet = header(wb, "Formats", "Call", "Result", "");

        row(sheet, 1, "Formats.format(1234.5678, \"$#,##0.00\")",
                Formats.format(1234.5678, "$#,##0.00"));
        row(sheet, 2, "Formats.format(0.0742, \"0.0%\")",
                Formats.format(0.0742, "0.0%"));
        row(sheet, 3, "Formats.format(46258, \"d mmm yyyy\")",
                Formats.format(46_258, "d mmm yyyy"));
        row(sheet, 4, "Formats.parse(\"1,234.56\")",
                describe(Formats.parse("1,234.56")));
        row(sheet, 5, "Formats.parse(\"1 234,56\")",
                describe(Formats.parse("1 234,56")));
        sheet.cell(7, 0).setValue(CellValue.text(
                "parse reports what it recognised. A string it cannot read as a number "
                        + "comes back as TEXT rather than as an error."));

        sheet.setColumnWidth(0, 42);
        sheet.setColumnWidth(1, 24);
        return wb;
    }

    /**
     * A parse result rendered as what it means.
     *
     * <p>{@link ParsedValue} is a tagged union over one record: {@code number} is
     * the payload for NUMBER, DATE and BOOL, and {@code text} is the payload for
     * TEXT and FORMULA, so only the field that carries the value is shown.
     */
    private static String describe(ParsedValue parsed) {
        switch (parsed.kind()) {
            case NUMBER:
            case DATE:
                return parsed.kind() + " " + parsed.number();
            case BOOL:
                return parsed.kind() + " " + parsed.booleanValue();
            default:
                return parsed.kind() + " \"" + parsed.text() + "\"";
        }
    }

    private static void row(Sheet sheet, int index, String call, String result) {
        sheet.cell(index, 0).setValue(CellValue.text(call));
        sheet.cell(index, 1).setValue(CellValue.text(result));
    }

    private static Sheet header(Workbook wb, String name, String a, String b, String c) {
        Sheet sheet = wb.createSheet(name);
        sheet.range("A1:C1").setValues(new CellValue[][] {{
                CellValue.text(a), CellValue.text(b), CellValue.text(c)}});
        sheet.range("A1:C1").setStyle(CellStyle.builder().bold().build());
        sheet.setColumnWidth(0, 26);
        sheet.setColumnWidth(1, 24);
        sheet.setColumnWidth(2, 20);
        return sheet;
    }
}
