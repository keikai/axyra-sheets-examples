package io.keikai.axyra.examples.featuredemo;

import io.keikai.axyra.examples.Example;
import io.keikai.axyra.sheets.CellValue;
import io.keikai.axyra.sheets.Range;
import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.Workbook;
import io.keikai.axyra.sheets.style.BorderStyle;
import io.keikai.axyra.sheets.style.CellStyle;
import io.keikai.axyra.sheets.style.Color;
import io.keikai.axyra.sheets.style.HAlign;
import io.keikai.axyra.sheets.style.VAlign;

import java.nio.file.Path;

/**
 * Merge cells and freeze panes — the two things that make a long sheet readable.
 *
 * <p>{@code merge()} joins a rectangle into one cell; {@code autoMerge()} does it
 * for you wherever consecutive cells repeat a value, which is what a grouped
 * report usually wants. {@code freezePanes(rows, cols)} pins the headings so they
 * stay put while the body scrolls.
 */
public final class MergeFreezeExample implements Example {
    @Override
    public String name() {
        return "feature-demo/merge-freeze";
    }

    @Override
    public String description() {
        return "Merge a title band, auto-merge repeated values and freeze panes";
    }

    @Override
    public void run(Path outputDirectory) {
        try (Workbook wb = titleBand()) {
            wb.save(outputDirectory.resolve("merge-freeze-title-band.xlsx"));
        }
        try (Workbook wb = autoMerge()) {
            wb.save(outputDirectory.resolve("merge-freeze-auto-merge.xlsx"));
        }
        try (Workbook wb = listMerges()) {
            wb.save(outputDirectory.resolve("merge-freeze-list-merges.xlsx"));
        }
        try (Workbook wb = freeze()) {
            wb.save(outputDirectory.resolve("merge-freeze-frozen.xlsx"));
        }
        System.out.println("Wrote four merge and freeze workbooks to " + outputDirectory);
    }

    /** The classic report header: one merged band across the columns. */
    private static Workbook titleBand() {
        Workbook wb = Workbook.create();
        Sheet sheet = ledger(wb, "Titled", 12);

        sheet.insertRows(0, 2);
        sheet.cell(0, 0).setValue(CellValue.text("Quarterly ledger — Q3 2026"));
        sheet.range("A1:D1").merge();
        sheet.range("A1:D1").setStyle(CellStyle.builder()
                .bold().fontSize(15).fontColor(Color.WHITE)
                .fill(Color.rgb(0x1F, 0x4E, 0x79))
                .horizontalAlign(HAlign.CENTER)
                .verticalAlign(VAlign.CENTER)
                .build());
        sheet.setRowHeight(0, 30);
        return wb;
    }

    /** Runs of equal values collapse into one merged cell each. */
    private static Workbook autoMerge() {
        Workbook wb = Workbook.create();
        Sheet sheet = ledger(wb, "Grouped", 12);

        int merged = sheet.range(1, 0, 12, 0).autoMerge(false);
        sheet.range(1, 0, 12, 0).setStyle(CellStyle.builder()
                .verticalAlign(VAlign.CENTER)
                .horizontalAlign(HAlign.CENTER)
                .allBorders(BorderStyle.THIN, Color.rgb(0xBF, 0xBF, 0xBF))
                .build());

        sheet.cell(14, 0).setValue(CellValue.text(
                "autoMerge(false) created " + merged + " vertical regions in column A."));
        sheet.setColumnWidth(0, 16);
        return wb;
    }

    /** Every merged rectangle is readable back as a {@link Range}. */
    private static Workbook listMerges() {
        Workbook wb = autoMerge();
        Sheet sheet = wb.sheet(0);

        int row = 16;
        sheet.cell(row, 0).setValue(CellValue.text("Merged regions"));
        sheet.range(row, 0, row, 0).setStyle(CellStyle.builder().bold().build());
        for (Range region : sheet.mergedRegions()) {
            row++;
            sheet.cell(row, 0).setValue(CellValue.text(
                    "r" + (region.firstRow() + 1) + "c" + (region.firstColumn() + 1)
                            + " to r" + (region.lastRow() + 1) + "c" + (region.lastColumn() + 1)));
        }
        return wb;
    }

    /** The heading row and the label column stay put while the rest scrolls. */
    private static Workbook freeze() {
        Workbook wb = Workbook.create();
        Sheet sheet = ledger(wb, "Frozen", 60);

        sheet.freezePanes(1, 1);
        sheet.cell(0, 4).setValue(CellValue.text(
                "freezePanes(1, 1) — scroll down, the header row stays."));
        sheet.setColumnWidth(4, 46);
        return wb;
    }

    /** A ledger whose first column repeats, so grouping has something to do. */
    private static Sheet ledger(Workbook wb, String name, int rows) {
        Sheet sheet = wb.createSheet(name);
        sheet.range("A1:D1").setValues(new CellValue[][] {{
                CellValue.text("Region"), CellValue.text("Month"),
                CellValue.text("Units"), CellValue.text("Revenue")}});
        sheet.range("A1:D1").setStyle(CellStyle.builder().bold().build());

        String[] regions = {"North", "South", "East", "West"};
        String[] months = {"Jul", "Aug", "Sep"};
        for (int i = 0; i < rows; i++) {
            sheet.cell(i + 1, 0).setValue(CellValue.text(regions[(i / 3) % regions.length]));
            sheet.cell(i + 1, 1).setValue(CellValue.text(months[i % months.length]));
            sheet.setNumber(i + 1, 2, 80 + (i * 37) % 220);
            sheet.setNumber(i + 1, 3, 1200 + (i * 613) % 4300);
        }
        sheet.range(1, 3, rows, 3).setStyle(CellStyle.builder().numberFormat("#,##0").build());
        sheet.setColumnWidth(0, 14);
        sheet.setColumnWidth(3, 14);
        return sheet;
    }
}
