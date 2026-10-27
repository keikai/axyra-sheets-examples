package io.keikai.axyra.examples.regression;

import io.keikai.axyra.sheets.AggregateFunction;
import io.keikai.axyra.sheets.AxyraException;
import io.keikai.axyra.sheets.PivotTableView;
import io.keikai.axyra.sheets.Range;
import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.SheetProtection;
import io.keikai.axyra.sheets.Visibility;
import io.keikai.axyra.sheets.Workbook;
import io.keikai.axyra.sheets.content.Chart;
import io.keikai.axyra.sheets.ChartView;
import io.keikai.axyra.sheets.content.DataLabels;
import io.keikai.axyra.sheets.content.LegendPosition;
import io.keikai.axyra.sheets.content.Trendline;
import io.keikai.axyra.sheets.io.ImageOptions;
import io.keikai.axyra.sheets.io.PdfOptions;
import io.keikai.axyra.sheets.style.CellStyle;
import io.keikai.axyra.sheets.style.Color;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.List;

import static io.keikai.axyra.examples.regression.IssueTestSupport.n;
import static io.keikai.axyra.examples.regression.IssueTestSupport.number;
import static io.keikai.axyra.examples.regression.IssueTestSupport.t;
import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * Features F24-F35 of the axyra-sheets skill test (charts, sparklines, pictures, shapes, pivot tables,
 * sheet management, protection, encryption, page setup, PDF and image output).
 *
 * <p>Saved files carry the evaluation notice sheet, so sheets are looked up by name and the notice
 * sheet is left out of sheet-order checks.
 */
class FeatureContentTest {
    private static final String NOTICE = "Axyra Evaluation Copy";

    /** Page objects in a PDF ({@code /Type /Page}, not {@code /Pages}). */
    private static int pages(String pdf) {
        return pdf.split("/Type\\s*/Page[^s]", -1).length - 1;
    }

    private static void quarterlyData(Sheet s) {
        String[] quarters = {"Q1", "Q2", "Q3", "Q4"};
        s.cell(0, 0).setValue(t("Quarter"));
        s.cell(0, 1).setValue(t("North"));
        s.cell(0, 2).setValue(t("South"));
        for (int r = 1; r <= 4; r++) {
            s.cell(r, 0).setValue(t(quarters[r - 1]));
            s.cell(r, 1).setValue(n(r * 100));
            s.cell(r, 2).setValue(n(r * 80));
        }
    }

    @Test
    @DisplayName("F24 a clustered column chart and a pie chart, reopened")
    void createCharts(@TempDir Path dir) {
        Path p = dir.resolve("charts.xlsx");
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("Sales");
            quarterlyData(s);
            s.addChart(Chart.column().title("Quarterly Sales")
                    .series("North", "Sales!$B$2:$B$5", "Sales!$A$2:$A$5")
                    .series("South", "Sales!$C$2:$C$5", "Sales!$A$2:$A$5")
                    .legend(LegendPosition.BOTTOM).categoryAxisTitle("Quarter").valueAxisTitle("Units")
                    .anchor(6, 0, 20, 8).build());
            s.addChart(Chart.pie().title("North Share").series("North", "Sales!$B$2:$B$5", "Sales!$A$2:$A$5")
                    .legend(LegendPosition.RIGHT).anchor(6, 9, 20, 16).build());
            wb.save(p);
        }
        try (Workbook wb = Workbook.open(p)) {
            ChartView column = wb.sheet("Sales").chart(0);
            ChartView pie = wb.sheet("Sales").chart(1);
            assertEquals("column", column.type());
            assertEquals("Quarterly Sales", column.title());
            assertEquals(2, column.seriesCount());
            assertEquals("clustered", column.grouping());
            assertEquals(LegendPosition.BOTTOM, column.legendPosition());
            assertEquals("pie", pie.type());
            assertEquals(1, pie.seriesCount());
        }
    }

    @Test
    @DisplayName("F25 change a line chart's title, fill, data labels and trendline, reopened")
    void modifyChart(@TempDir Path dir) {
        Path p = dir.resolve("line.xlsx");
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("Sales");
            quarterlyData(s);
            int i = s.addChart(Chart.line().title("Sales").series("North", "Sales!$B$2:$B$5", "Sales!$A$2:$A$5")
                    .anchor(6, 0, 20, 8).build());
            s.chart(i).setTitle("Sales (Revised)")
                    .setSeriesFill(0, Color.rgb(0x1F, 0x77, 0xB4))
                    .setShowDataLabels(true)
                    .setSeriesDataLabels(0, DataLabels.value())
                    .addTrendline(0, Trendline.linear());
            wb.save(p);
        }
        try (Workbook wb = Workbook.open(p)) {
            ChartView v = wb.sheet("Sales").chart(0);
            assertEquals("Sales (Revised)", v.title());
            assertEquals("FF1F77B4", v.seriesFill(0).toHex().replace("#", ""));
            assertTrue(v.seriesDataLabels(0).showValue());
            assertEquals(1, v.trendlines(0).size());
        }
    }

    @Test
    @DisplayName("F26 one sparkline per data row, reopened")
    void sparklines(@TempDir Path dir) {
        Path p = dir.resolve("sparklines.xlsx");
        try (Workbook wb = Workbook.create()) {
            Sheet d = wb.createSheet("Data");
            for (int r = 1; r <= 5; r++) {
                for (int c = 1; c <= 12; c++) {
                    d.cell(r, c).setValue(n((r * c) % 7));
                }
            }
            d.addSparkline("line", "Data!B2:M6", "N2:N6");
            wb.save(p);
        }
        try (Workbook wb = Workbook.open(p)) {
            var groups = wb.sheet("Data").sparklines();
            assertEquals(1, groups.size());
            assertEquals("line", groups.get(0).type());
            assertEquals(5, groups.get(0).sparklines().size());
        }
    }

    @Test
    @DisplayName("F27 a generated PNG anchored at a cell, reopened, then removed")
    void pictures(@TempDir Path dir) throws Exception {
        BufferedImage image = new BufferedImage(40, 20, BufferedImage.TYPE_INT_RGB);
        image.createGraphics().fillRect(0, 0, 20, 20);
        ByteArrayOutputStream png = new ByteArrayOutputStream();
        ImageIO.write(image, "png", png);
        Path p = dir.resolve("picture.xlsx");
        try (Workbook wb = Workbook.create()) {
            wb.createSheet("S").addPicture(png.toByteArray(), "png", 1, 1, 6, 5);
            wb.save(p);
        }
        try (Workbook wb = Workbook.open(p)) {
            Sheet s = wb.sheet("S");
            assertEquals(1, s.pictureCount());
            assertArrayEquals(new int[] {1, 1, 6, 5}, s.pictureAnchor(0));
            assertEquals("png", s.pictureFormat(0));
            assertTrue(s.removePicture(0));
            assertEquals(0, s.pictureCount());
        }
    }

    @Test
    @DisplayName("F28 a filled rectangle with text and a text box, reopened")
    void shapes(@TempDir Path dir) {
        Path p = dir.resolve("shapes.xlsx");
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("S");
            int rect = s.addShape("rect", "Review", 4, 7, 7, 11);
            s.setShapeFillColor(rect, Color.rgb(255, 204, 0));
            s.addTextBox("Quarterly sales", 0, 7, 3, 11);
            wb.save(p);
        }
        try (Workbook wb = Workbook.open(p)) {
            Sheet s = wb.sheet("S");
            assertEquals(2, s.shapeCount());
            assertEquals("rect", s.shapeType(0));
            assertEquals("Review", s.shapeText(0));
            assertEquals("FFFFCC00", s.shapeFillColor(0).toHex().replace("#", ""));
            assertTrue(s.isTextBox(1));
            assertEquals("Quarterly sales", s.shapeText(1));
        }
    }

    @Test
    @DisplayName("F29 a pivot of region by product with the sum of amount")
    void pivotTable() {
        try (Workbook wb = Workbook.create()) {
            Sheet d = wb.createSheet("Data");
            String[][] rows = {{"Region", "Product", "Amount"}, {"North", "Apple", "150"}, {"North", "Banana", "80"},
                    {"South", "Apple", "200"}, {"South", "Banana", "90"}};
            for (int r = 0; r < rows.length; r++) {
                for (int c = 0; c < 3; c++) {
                    d.cell(r, c).setValue(r > 0 && c == 2 ? n(Double.parseDouble(rows[r][c])) : t(rows[r][c]));
                }
            }
            Sheet ps = wb.createSheet("Pivot");
            PivotTableView pv = ps.pivotTable(ps.addPivotTable("SalesPivot", "A1", "Data!$A$1:$C$5"));
            pv.addRowField("Region");
            pv.addColumnField("Product");
            pv.addDataField("Amount", AggregateFunction.SUM);
            pv.refresh();
            Range used = ps.usedRange();
            int grandRow = -1;
            for (int r = used.firstRow(); r <= used.lastRow(); r++) {
                if (t("Grand Total").equals(ps.cell(r, 0).value())) {
                    grandRow = r;
                }
            }
            assertTrue(grandRow > 0, "no Grand Total row");
            assertEquals(520.0, number(ps.cell(grandRow, used.lastColumn()).value()));
        }
    }

    @Test
    @DisplayName("F30 rename, move, copy, delete, hide, tab colour and active sheet, reopened")
    void sheetManagement(@TempDir Path dir) {
        Path p = dir.resolve("sheets.xlsx");
        try (Workbook wb = Workbook.create()) {
            for (String name : new String[] {"Data", "Notes", "Summary", "Temp", "Archive"}) {
                wb.createSheet(name).cell(0, 0).setValue(t(name));
            }
            wb.renameSheet(wb.sheet("Notes").index(), "Readme");
            wb.moveSheet(wb.sheet("Summary").index(), 0);
            wb.copySheet(wb.sheet("Data"), "Data Copy");
            wb.deleteSheet(wb.sheet("Temp").index());
            wb.sheet("Archive").setVisibility(Visibility.HIDDEN);
            wb.sheet("Summary").setTabColor(Color.rgb(255, 128, 0));
            wb.setActiveSheet(wb.sheet("Summary").index());
            wb.save(p);
        }
        try (Workbook wb = Workbook.open(p)) {
            List<String> names = wb.sheets().stream().map(Sheet::name).filter(x -> !x.startsWith(NOTICE)).toList();
            assertEquals(List.of("Summary", "Data", "Readme", "Archive", "Data Copy"), names);
            assertEquals(Visibility.HIDDEN, wb.sheet("Archive").visibility());
            assertEquals("FFFF8000", wb.sheet("Summary").tabColor().toHex().replace("#", ""));
            assertEquals(t("Data"), wb.sheet("Data Copy").cell(0, 0).value());
        }
    }

    @Test
    @DisplayName("F31 format-only sheet protection, unlocked cells, structure and a protected range")
    void protection(@TempDir Path dir) {
        Path p = dir.resolve("protected.xlsx");
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("Data");
            s.cell(1, 1).setStyle(CellStyle.builder().locked(false).build());
            s.protect("secret", SheetProtection.builder().formatCells(true).formatColumns(true).formatRows(true).build());
            s.addProtectedRange("Totals", "D2:D5", "range-pw");
            wb.protectStructure("book-pw");
            wb.save(p);
        }
        try (Workbook wb = Workbook.open(p)) {
            Sheet s = wb.sheet("Data");
            assertTrue(s.isProtected());
            assertTrue(s.protection().formatCells());
            assertFalse(s.cell(1, 1).style().locked());
            assertTrue(s.cell(0, 0).style().locked());
            assertEquals("Totals", s.protectedRanges().get(0).name());
            assertTrue(wb.isStructureProtected());
            assertTrue(s.unprotect("secret"));
            assertFalse(s.isProtected());
            assertTrue(wb.unprotectStructure("book-pw"));
        }
    }

    @Test
    @DisplayName("F32 an encrypted workbook reopens with its password; write protection verifies")
    void encryptionAndWriteProtection(@TempDir Path dir) {
        Path encrypted = dir.resolve("encrypted.xlsx");
        Path writeProtected = dir.resolve("write-protected.xlsx");
        try (Workbook wb = Workbook.create()) {
            wb.createSheet("S").cell(0, 0).setValue(t("Secret"));
            wb.saveEncrypted(encrypted, "pw");
            wb.setWriteProtection("modify-pw", "Author", true);
            assertTrue(wb.verifyWritePassword("modify-pw"));
            assertFalse(wb.verifyWritePassword("wrong"));
            wb.save(writeProtected);
        }
        try (Workbook wb = Workbook.open(encrypted, "pw")) {
            assertEquals(t("Secret"), wb.sheet("S").cell(0, 0).value());
        }
        assertThrows(AxyraException.class, () -> Workbook.open(encrypted).close());
        try (Workbook wb = Workbook.open(writeProtected)) {
            assertTrue(wb.writeProtection().readOnlyRecommended());
            assertEquals("Author", wb.writeProtection().userName());
            assertTrue(wb.verifyWritePassword("modify-pw"));
        }
    }

    @Test
    @DisplayName("F33 page setup with a manual row page break, reopened")
    void pageSetup(@TempDir Path dir) {
        Path p = dir.resolve("page.xlsx");
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("Report");
            for (int r = 0; r < 80; r++) {
                s.cell(r, 0).setValue(n(r));
            }
            s.setLandscape(true);
            s.setPaperSize(9);
            s.setPageMargins(0.5, 0.5, 0.75, 0.75, 0.3, 0.3);
            s.setPrintArea("A1:F81");
            s.setPrintTitleRows("$1:$1");
            s.fitToPage(1, 0);
            s.setHeader("&CQuarterly Sales Report");
            s.setFooter("&RPage &P of &N");
            s.addRowPageBreak(40);
            wb.save(p);
        }
        try (Workbook wb = Workbook.open(p)) {
            Sheet s = wb.sheet("Report");
            assertTrue(s.isLandscape());
            assertEquals(9, s.paperSize());
            assertEquals(0.75, s.pageMargins().top(), 1e-9);
            assertEquals("A1:F81", s.printArea());
            assertEquals("$1:$1", s.printTitleRows());
            // the evaluation build appends its notice to the header
            assertTrue(s.header().startsWith("&CQuarterly Sales Report"), s.header());
            assertArrayEquals(new int[] {40}, s.rowPageBreaks());
        }
    }

    @Test
    @DisplayName("F34 PDF export fitted to width, and a page range")
    void pdfExport(@TempDir Path dir) throws Exception {
        Path all = dir.resolve("out.pdf");
        Path first = dir.resolve("first.pdf");
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("Report");
            for (int r = 0; r < 200; r++) {
                s.cell(r, 0).setValue(t("Row " + r));
                s.cell(r, 1).setValue(n(r));
            }
            wb.renderPdf(all, PdfOptions.builder().sheetNames("Report").fitTo(PdfOptions.Fit.WIDTH).build());
            wb.renderPdf(first, PdfOptions.builder().sheetNames("Report").pageRange("1").build());
        }
        String pdf = new String(Files.readAllBytes(all), StandardCharsets.ISO_8859_1);
        assertTrue(pdf.startsWith("%PDF"));
        assertTrue(pages(pdf) > 1, "expected several pages, got " + pages(pdf));
        String one = new String(Files.readAllBytes(first), StandardCharsets.ISO_8859_1);
        assertEquals(1, pages(one));
    }

    @Test
    @DisplayName("F35 render a sheet to PNG at 2x and to SVG")
    void renderImages() {
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("Report");
            for (int r = 0; r < 10; r++) {
                s.cell(r, 0).setValue(n(r));
            }
            int index = s.index();
            byte[] png = wb.renderSheetPng(index, ImageOptions.builder().scale(2.0).build());
            byte[] svg = wb.renderSheetSvg(index, ImageOptions.builder().build());
            assertArrayEquals(new byte[] {(byte) 0x89, 'P', 'N', 'G'}, Arrays.copyOf(png, 4));
            int width = ((png[16] & 0xFF) << 24) | ((png[17] & 0xFF) << 16) | ((png[18] & 0xFF) << 8) | (png[19] & 0xFF);
            assertTrue(width > 1500 && width < 1700, "A4 at 2x is about 1588 px wide, was " + width);
            assertTrue(new String(svg, StandardCharsets.UTF_8).contains("<svg"));
            assertEquals(1, wb.sheetPageCount(index));
        }
    }
}
