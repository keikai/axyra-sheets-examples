package io.keikai.axyra.examples.featuredemo;

import io.keikai.axyra.examples.Example;
import io.keikai.axyra.sheets.CellValue;
import io.keikai.axyra.sheets.Range;
import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.Workbook;
import io.keikai.axyra.sheets.style.CellStyle;
import io.keikai.axyra.sheets.style.Color;

import java.nio.file.Path;

/**
 * Find and replace.
 *
 * <p>{@code find} returns <b>every</b> hit as {@code int[][]} of row/column
 * pairs — there is no cursor to advance and no "find next" to loop. Options
 * cover case sensitivity, whole-cell matching and {@code *}/{@code ?} wildcards,
 * and there is a regular-expression pair alongside for the harder patterns.
 */
public final class FindReplaceExample implements Example {
    private static final String CONTENT = "findReplace-tickets-content.xlsx";

    @Override
    public String name() {
        return "feature-demo/find-replace";
    }

    @Override
    public String description() {
        return "Find every match with options or a regex, and replace in one call";
    }

    @Override
    public void run(Path outputDirectory) {
        try (Workbook wb = findAndHighlight()) {
            wb.save(outputDirectory.resolve("find-replace-highlight.xlsx"));
        }
        try (Workbook wb = findWithOptions()) {
            wb.save(outputDirectory.resolve("find-replace-options.xlsx"));
        }
        try (Workbook wb = replaceAll()) {
            wb.save(outputDirectory.resolve("find-replace-replace-all.xlsx"));
        }
        try (Workbook wb = regexReplace()) {
            wb.save(outputDirectory.resolve("find-replace-regex.xlsx"));
        }
        System.out.println("Wrote four find and replace workbooks to " + outputDirectory);
    }

    /** Every hit comes back at once, so highlighting them is a single loop. */
    private static Workbook findAndHighlight() {
        Workbook wb = tickets();
        Sheet sheet = wb.sheet(0);

        int[][] hits = sheet.find("overdue");
        CellStyle flag = CellStyle.builder()
                .fill(Color.rgb(0xFF, 0xC7, 0xCE))
                .fontColor(Color.rgb(0x9C, 0x00, 0x06))
                .bold().build();
        for (int[] hit : hits) {
            sheet.range(hit[0], hit[1], hit[0], hit[1]).setStyle(flag);
        }

        note(sheet, "find(\"overdue\") returned " + hits.length + " hits — all of them, "
                + "in one call.");
        return wb;
    }

    /** Case, whole-cell and wildcard switches, counted side by side. */
    private static Workbook findWithOptions() {
        Workbook wb = tickets();
        Sheet sheet = wb.sheet(0);

        int base = sheet.usedRange().lastRow() + 3;
        sheet.range(base, 0, base, 1).setValues(new CellValue[][] {{
                CellValue.text("Call"), CellValue.text("Hits")}});
        sheet.range(base, 0, base, 1).setStyle(CellStyle.builder().bold().build());

        count(sheet, base + 1, "find(\"overdue\")",
                sheet.find("overdue").length);
        count(sheet, base + 2, "find(\"OVERDUE\", caseSensitive=true, ...)",
                sheet.find("OVERDUE", true, false, false).length);
        count(sheet, base + 3, "find(\"overdue\", ..., matchEntire=true, ...)",
                sheet.find("overdue", false, true, false).length);
        count(sheet, base + 4, "find(\"over*\", ..., wildcard=true)",
                sheet.find("over*", false, false, true).length);
        count(sheet, base + 5, "findRegex(\"^INV-\\\\d{4}$\")",
                sheet.findRegex("^INV-\\d{4}$").length);

        sheet.setColumnWidth(0, 42);
        return wb;
    }

    /** {@code replaceAll} returns how many cells it changed. */
    private static Workbook replaceAll() {
        Workbook wb = tickets();
        Sheet sheet = wb.sheet(0);

        int changed = sheet.replaceAll("overdue", "ESCALATED");
        note(sheet, "replaceAll(\"overdue\", \"ESCALATED\") changed " + changed + " cells.");

        CellStyle flag = CellStyle.builder()
                .fill(Color.rgb(0xFF, 0xEB, 0x9C))
                .fontColor(Color.rgb(0x9C, 0x65, 0x00)).bold().build();
        for (int[] hit : sheet.find("ESCALATED")) {
            sheet.range(hit[0], hit[1], hit[0], hit[1]).setStyle(flag);
        }
        return wb;
    }

    /** Regular expressions reach the patterns wildcards cannot. */
    private static Workbook regexReplace() {
        Workbook wb = tickets();
        Sheet sheet = wb.sheet(0);

        // Re-number every invoice from INV-1234 to 2026/1234.
        int changed = sheet.replaceAllRegex("INV-(\\d{4})", "2026/$1");
        note(sheet, "replaceAllRegex(\"INV-(\\\\d{4})\", \"2026/$1\") rewrote "
                + changed + " references.");
        return wb;
    }

    private static void count(Sheet sheet, int row, String call, int hits) {
        sheet.cell(row, 0).setValue(CellValue.text(call));
        sheet.setNumber(row, 1, hits);
    }

    private static void note(Sheet sheet, String text) {
        int row = sheet.usedRange().lastRow() + 2;
        sheet.cell(row, 0).setValue(CellValue.text(text));
        sheet.range(row, 0, row, 0).setStyle(CellStyle.builder().italic().build());
    }

    /**
     * The ticket list, read from a packaged workbook.
     *
     * <p>The used block is read and written straight back: {@code find} and
     * {@code replaceAll} match only cells that have been written through the Java
     * API, so cells that arrived through the importer are invisible to them until
     * something rewrites them. The rewrite is value-for-value, so nothing about
     * the sheet changes but its provenance.
     */
    private static Workbook tickets() {
        Workbook wb = Workbook.openBytes(DemoBooks.read(CONTENT), "xlsx");
        Sheet sheet = wb.sheet(0);
        Range used = sheet.usedRange();
        sheet.range(used.firstRow(), used.firstColumn(), used.lastRow(), used.lastColumn())
                .setValues(used.values());
        return wb;
    }
}
