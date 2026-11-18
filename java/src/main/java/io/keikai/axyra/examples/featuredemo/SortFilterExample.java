package io.keikai.axyra.examples.featuredemo;

import io.keikai.axyra.examples.Example;
import io.keikai.axyra.sheets.CellValue;
import io.keikai.axyra.sheets.FilterOperator;
import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.SortKey;
import io.keikai.axyra.sheets.Workbook;
import io.keikai.axyra.sheets.style.CellStyle;

import java.nio.file.Path;

/**
 * Sort and filter.
 *
 * <p>Sorting rewrites the rows in place; {@link SortKey} stacks as many keys as
 * you need. Filtering hides rows rather than removing them —
 * {@code applyAutoFilter()} returns how many it hid, and the data is all still
 * there.
 *
 * <p>AutoFilter custom criteria compare <b>exactly</b>. For begins-with or
 * contains, use {@code applyAdvancedFilter}, which does honour {@code *} and
 * {@code ?}.
 */
public final class SortFilterExample implements Example {
    private static final String CONTENT = "sortFilter-orders-content.xlsx";
    private static final int ROWS = 24;

    @Override
    public String name() {
        return "feature-demo/sort-filter";
    }

    @Override
    public String description() {
        return "Sort by one or several keys and hide rows with AutoFilter criteria";
    }

    @Override
    public void run(Path outputDirectory) {
        try (Workbook wb = unsorted()) {
            wb.save(outputDirectory.resolve("sort-filter-unsorted.xlsx"));
        }
        try (Workbook wb = sortOne()) {
            wb.save(outputDirectory.resolve("sort-filter-sort-one.xlsx"));
        }
        try (Workbook wb = sortTwo()) {
            wb.save(outputDirectory.resolve("sort-filter-sort-two.xlsx"));
        }
        try (Workbook wb = filterValue()) {
            wb.save(outputDirectory.resolve("sort-filter-filter-value.xlsx"));
        }
        try (Workbook wb = filterComparison()) {
            wb.save(outputDirectory.resolve("sort-filter-filter-comparison.xlsx"));
        }
        System.out.println("Wrote five sort and filter workbooks to " + outputDirectory);
    }

    /** The source data, in the order it was written. */
    private static Workbook unsorted() {
        Workbook wb = orders();
        Sheet sheet = wb.sheet(0);
        note(sheet, "Straight from the writer — regions interleaved, revenue unordered.");
        return wb;
    }

    /** Revenue, highest first. */
    private static Workbook sortOne() {
        Workbook wb = orders();
        Sheet sheet = wb.sheet(0);
        sheet.range(1, 0, ROWS, 3).sort(3, false);
        note(sheet, "sort(column 3, ascending=false) — highest revenue first.");
        return wb;
    }

    /** Region ascending, then revenue descending inside each region. */
    private static Workbook sortTwo() {
        Workbook wb = orders();
        Sheet sheet = wb.sheet(0);
        sheet.range(1, 0, ROWS, 3).sort(SortKey.asc(0), SortKey.desc(3));
        note(sheet, "sort(SortKey.asc(0), SortKey.desc(3)) — grouped, then ranked.");
        return wb;
    }

    /** Show only the rows whose region matches. */
    private static Workbook filterValue() {
        Workbook wb = orders();
        Sheet sheet = wb.sheet(0);

        sheet.setAutoFilter(sheet.range(0, 0, ROWS, 3));
        sheet.setColumnFilter(0, "East");
        int hidden = sheet.applyAutoFilter();

        note(sheet, "setColumnFilter(0, \"East\") hid " + hidden + " of " + ROWS
                + " rows. The data is intact — only the row flags changed.");
        return wb;
    }

    /** Custom criteria compare exactly; here that is what we want. */
    private static Workbook filterComparison() {
        Workbook wb = orders();
        Sheet sheet = wb.sheet(0);

        sheet.setAutoFilter(sheet.range(0, 0, ROWS, 3));
        sheet.setColumnCustomFilter(3, FilterOperator.GREATER_THAN, "3000");
        int hidden = sheet.applyAutoFilter();

        note(sheet, "setColumnCustomFilter(3, GREATER_THAN, \"3000\") hid "
                + hidden + " rows. Wildcards are not supported here — "
                + "applyAdvancedFilter handles those.");
        return wb;
    }

    private static void note(Sheet sheet, String text) {
        int row = ROWS + 2;
        sheet.cell(row, 0).setValue(CellValue.text(text));
        sheet.range(row, 0, row, 0).setStyle(CellStyle.builder().italic().build());
    }

    /** The order book: deliberately unsorted, with repeated regions, read from a packaged workbook. */
    private static Workbook orders() {
        return Workbook.openBytes(DemoBooks.read(CONTENT), "xlsx");
    }
}
