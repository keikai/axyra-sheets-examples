package io.keikai.axyra.examples.featuredemo;

import io.keikai.axyra.examples.Example;
import io.keikai.axyra.sheets.CellValue;
import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.Workbook;
import io.keikai.axyra.sheets.content.Chart;
import io.keikai.axyra.sheets.content.LegendPosition;
import io.keikai.axyra.sheets.io.PdfOptions;
import io.keikai.axyra.sheets.io.Renderer;
import io.keikai.axyra.sheets.style.CellStyle;
import io.keikai.axyra.sheets.style.Color;
import io.keikai.axyra.sheets.style.HAlign;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;

/**
 * Export to PDF.
 *
 * <p>{@link Renderer#toPdf} returns the file as a {@code byte[]}, which is what a
 * server-side report actually needs — straight into a download, an email
 * attachment or a blob column, with no temporary file. {@link PdfOptions}
 * selects sheets, a printed-page range, and how the content is fitted;
 * {@code sheetPageCount} says in advance how many pages it will be.
 */
public final class ExportPdfExample implements Example {
    @Override
    public String name() {
        return "feature-demo/export-pdf";
    }

    @Override
    public String description() {
        return "Render a workbook to PDF bytes: whole book, one sheet, fitted, and page counts";
    }

    @Override
    public void run(Path outputDirectory) throws IOException {
        try (Workbook wb = report()) {
            wb.save(outputDirectory.resolve("export-pdf-report.xlsx"));
        }
        write(outputDirectory.resolve("report.pdf"), whole());
        write(outputDirectory.resolve("report-sheet1.pdf"), firstSheet());
        write(outputDirectory.resolve("report-fit-width.pdf"), fitted());
        try (Workbook wb = pageCount()) {
            wb.save(outputDirectory.resolve("export-pdf-page-count.xlsx"));
        }
    }

    private static void write(Path file, byte[] pdf) throws IOException {
        Files.write(file, pdf);
        System.out.println("Wrote " + file.getFileName() + " (" + pdf.length + " bytes)");
    }

    /** Every sheet, with the default options. */
    private static byte[] whole() {
        try (Workbook wb = report()) {
            return Renderer.toPdf(wb, PdfOptions.builder().build());
        }
    }

    /** {@code sheets(...)} takes indexes; {@code sheetNames(...)} takes names. */
    private static byte[] firstSheet() {
        try (Workbook wb = report()) {
            return Renderer.toPdf(wb, PdfOptions.builder().sheets(0).build());
        }
    }

    /** Squeeze a wide sheet onto one page across. */
    private static byte[] fitted() {
        try (Workbook wb = report()) {
            return Renderer.toPdf(wb, PdfOptions.builder()
                    .fitTo(PdfOptions.Fit.WIDTH)
                    .build());
        }
    }

    /** How long the PDF will be, before rendering it. */
    private static Workbook pageCount() {
        Workbook wb = report();
        Sheet sheet = wb.sheet(0);
        int row = 30;   // clear of the chart, which occupies rows 12-29

        sheet.cell(row, 0).setValue(CellValue.text("Pages — Summary"));
        sheet.setNumber(row, 3, wb.sheetPageCount(0));
        sheet.cell(row + 1, 0).setValue(CellValue.text("Pages — By rep"));
        sheet.setNumber(row + 1, 3, wb.sheetPageCount(1));
        sheet.cell(row + 2, 0).setValue(CellValue.text("Renderer.pageCount(wb, 0)"));
        sheet.setNumber(row + 2, 3, Renderer.pageCount(wb, 0));
        sheet.cell(row + 3, 0).setValue(CellValue.text(
                "Whole workbook prints both; sheets(0) prints the first alone."));
        sheet.range(row, 0, row + 2, 0).setStyle(CellStyle.builder().bold().build());
        sheet.setColumnWidth(0, 30);
        return wb;
    }

    /** A report worth printing: a title band, figures, a total and a chart. */
    private static Workbook report() {
        Workbook wb = Workbook.create();
        Sheet sheet = wb.createSheet("Summary");

        sheet.cell(0, 0).setValue(CellValue.text("Regional performance — Q3 2026"));
        sheet.range("A1:D1").merge();
        sheet.range("A1:D1").setStyle(CellStyle.builder()
                .bold().fontSize(15).fontColor(Color.WHITE)
                .fill(Color.rgb(0x1F, 0x4E, 0x79))
                .horizontalAlign(HAlign.CENTER).build());
        sheet.setRowHeight(0, 30);

        sheet.range("A3:D3").setValues(new CellValue[][] {{
                CellValue.text("Region"), CellValue.text("Units"),
                CellValue.text("Revenue"), CellValue.text("Margin")}});
        sheet.range("A3:D3").setStyle(CellStyle.builder().bold().build());

        String[] regions = {"North", "South", "East", "West", "Central"};
        double[] units = {1240, 980, 1610, 1130, 870};
        double[] revenue = {18600, 14700, 24150, 16950, 13050};
        for (int i = 0; i < regions.length; i++) {
            sheet.cell(i + 3, 0).setValue(CellValue.text(regions[i]));
            sheet.setNumber(i + 3, 1, units[i]);
            sheet.setNumber(i + 3, 2, revenue[i]);
            sheet.cell(i + 3, 3).setFormula("C" + (i + 4) + "/B" + (i + 4));
        }

        int totalRow = regions.length + 3;
        sheet.cell(totalRow, 0).setValue(CellValue.text("Total"));
        sheet.cell(totalRow, 1).setFormula("SUM(B4:B8)");
        sheet.cell(totalRow, 2).setFormula("SUM(C4:C8)");
        sheet.range(totalRow, 0, totalRow, 3).setStyle(CellStyle.builder().bold().build());
        sheet.range(3, 2, totalRow, 2).setStyle(
                CellStyle.builder().numberFormat("#,##0").build());
        sheet.range(3, 3, totalRow - 1, 3).setStyle(
                CellStyle.builder().numberFormat("#,##0.00").build());
        sheet.setColumnWidth(0, 16);
        sheet.setColumnWidth(2, 14);

        sheet.addChart(Chart.column()
                .title("Revenue by region")
                .series("Revenue", "Summary!C4:C8", "Summary!A4:A8")
                .legend(LegendPosition.NONE)
                .anchor(11, 0, 28, 6)
                .build());

        // The summary prints the table and the chart below it, which ends in row 29.
        sheet.setPrintArea("A1:G29");

        detail(wb);
        wb.recalculate();
        return wb;
    }

    /**
     * A second sheet, long enough to be a page of its own, so that
     * {@code PdfOptions.sheets} has something to select.
     */
    private static void detail(Workbook wb) {
        Sheet sheet = wb.createSheet("By rep");

        sheet.cell(0, 0).setValue(CellValue.text("Revenue by representative — Q3 2026"));
        sheet.range("A1:C1").merge();
        sheet.range("A1:C1").setStyle(CellStyle.builder()
                .bold().fontSize(15).fontColor(Color.WHITE)
                .fill(Color.rgb(0x1F, 0x4E, 0x79))
                .horizontalAlign(HAlign.CENTER).build());
        sheet.setRowHeight(0, 30);

        sheet.range("A3:C3").setValues(new CellValue[][] {{
                CellValue.text("Rep"), CellValue.text("Region"), CellValue.text("Revenue")}});
        sheet.range("A3:C3").setStyle(CellStyle.builder().bold().build());

        String[] reps = {"Adler", "Brandt", "Cortez", "Duval", "Eriksen",
                "Fischer", "Goddard", "Halvorsen", "Ibarra", "Jansen"};
        String[] regions = {"North", "South", "East", "West", "Central"};
        int firstData = 3;
        for (int i = 0; i < reps.length; i++) {
            sheet.cell(firstData + i, 0).setValue(CellValue.text(reps[i]));
            sheet.cell(firstData + i, 1).setValue(CellValue.text(regions[i % regions.length]));
            sheet.setNumber(firstData + i, 2, 2600 + (i * 737) % 4200);
        }

        int lastData = firstData + reps.length - 1;
        int totalRow = lastData + 2;
        sheet.cell(totalRow, 0).setValue(CellValue.text("Total"));
        sheet.cell(totalRow, 2).setFormula(
                "SUM(C" + (firstData + 1) + ":C" + (lastData + 1) + ")");
        sheet.range(firstData, 2, totalRow, 2).setStyle(
                CellStyle.builder().numberFormat("#,##0").build());
        sheet.range(totalRow, 0, totalRow, 2).setStyle(CellStyle.builder()
                .bold().numberFormat("#,##0").build());
        sheet.setColumnWidth(0, 18);
        sheet.setColumnWidth(1, 14);
        sheet.setColumnWidth(2, 14);
        sheet.setPrintArea("A1:C" + (totalRow + 1));
    }
}
