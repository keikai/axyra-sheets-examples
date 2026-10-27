package io.keikai.axyra.examples.regression;

import io.keikai.axyra.sheets.AggregateFunction;
import io.keikai.axyra.sheets.AxyraFormulaException;
import io.keikai.axyra.sheets.CellValue;
import io.keikai.axyra.sheets.ConsolidateFunction;
import io.keikai.axyra.sheets.License;
import io.keikai.axyra.sheets.PivotTableView;
import io.keikai.axyra.sheets.Range;
import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.Workbook;
import io.keikai.axyra.sheets.content.Chart;
import io.keikai.axyra.sheets.content.ConditionalFormat;
import io.keikai.axyra.sheets.io.PdfOptions;
import io.keikai.axyra.sheets.io.Renderer;
import io.keikai.axyra.sheets.io.StreamWorkbook;
import io.keikai.axyra.sheets.style.CellStyle;
import io.keikai.axyra.sheets.style.Color;
import io.keikai.axyra.sheets.style.Theme;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Tag;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Base64;
import java.util.HashSet;
import java.util.List;

import static io.keikai.axyra.examples.regression.IssueTestSupport.count;
import static io.keikai.axyra.examples.regression.IssueTestSupport.findAll;
import static io.keikai.axyra.examples.regression.IssueTestSupport.n;
import static io.keikai.axyra.examples.regression.IssueTestSupport.number;
import static io.keikai.axyra.examples.regression.IssueTestSupport.pdfStreams;
import static io.keikai.axyra.examples.regression.IssueTestSupport.sheetsXml;
import static io.keikai.axyra.examples.regression.IssueTestSupport.t;
import static io.keikai.axyra.examples.regression.IssueTestSupport.zipEntry;
import static io.keikai.axyra.examples.regression.IssueTestSupport.zipNames;
import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;
import static org.junit.jupiter.api.Assumptions.assumeFalse;

/**
 * Open Axyra Sheets defects (items 35-61 of the 1021 benchmark issue list, still present in the build
 * this project uses). Each test asserts the behaviour Excel has, so it fails until Axyra fixes the
 * defect.
 *
 * <p>Tagged {@code known-issue} and left out of {@code ./mvnw test}. Run them with
 * {@code ./mvnw test -Pknown-issues}; a test that starts to pass means the defect is fixed and the test
 * can move to {@link FixedIssuesTest}.
 */
@Tag("known-issue")
class KnownIssuesTest {

    // ---- P1: fix before the next release ----

    @Test
    @DisplayName("#35 P1 PdfOptions.watermark is drawn after the page background")
    void watermarkIsNotPaintedOver() {
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("S");
            for (int r = 0; r < 30; r++) {
                s.cell(r, 0).setValue(t("Row " + r));
                s.cell(r, 1).setValue(n(r));
            }
            byte[] pdf = Renderer.toPdf(wb, PdfOptions.builder().watermark("CONFIDENTIAL").build());
            String page = pdfStreams(pdf).stream()
                    .filter(c -> c.contains("BT") && c.contains("1 1 1 rg"))
                    .findFirst().orElseThrow();
            assertTrue(page.indexOf("1 1 1 rg") < page.indexOf("BT"),
                    "the full-page white fill is painted after the first text (the watermark)");
        }
    }

    @Test
    @DisplayName("#36 P1 evaluation mode: re-saving a file adds no second notice sheet")
    void evaluationNoticeIsNotDuplicated(@TempDir Path dir) throws Exception {
        assumeFalse(License.status().isLicensed(), "needs evaluation mode");
        Path a = dir.resolve("a.xlsx");
        Path b = dir.resolve("b.xlsx");
        try (Workbook wb = Workbook.create()) {
            wb.createSheet("Q3").cell(0, 0).setValue(t("Region"));
            wb.save(a);
        }
        try (Workbook wb = Workbook.open(a)) {
            wb.save(b);
        }
        try (Workbook wb = Workbook.open(b)) {
            assertEquals(2, wb.sheetCount(), "one notice sheet plus Q3");
        }
    }

    @Test
    @DisplayName("#37 P1 an unterminated range in setFormula is a formula error, not an engine panic")
    void unterminatedRangeIsAFormulaError() {
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("S");
            assertThrows(AxyraFormulaException.class, () -> s.cell(0, 0).setFormula("SUM(B2:"));
        }
    }

    @Test
    @DisplayName("#38 P1 containsText, beginsWith and endsWith rules are saved with their own types")
    void textRulesKeepTheirType(@TempDir Path dir) throws Exception {
        Path p = dir.resolve("text-rules.xlsx");
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("S");
            String[] values = {"apple pie", "banana", "pineapple", "grape"};
            for (int r = 0; r < values.length; r++) {
                s.cell(r, 0).setValue(t(values[r]));
            }
            s.range("A1:A4").addConditionalFormat(ConditionalFormat.builder().containsText("apple").fillColor(Color.parse("#FFC7CE")).build());
            s.range("A1:A4").addConditionalFormat(ConditionalFormat.builder().beginsWith("pine").fillColor(Color.parse("#C6EFCE")).build());
            s.range("A1:A4").addConditionalFormat(ConditionalFormat.builder().endsWith("pe").fillColor(Color.parse("#FFEB9C")).build());
            wb.save(p);
        }
        String xml = sheetsXml(p);
        assertTrue(xml.contains("type=\"containsText\""), "no containsText rule");
        assertTrue(xml.contains("type=\"beginsWith\""), "no beginsWith rule");
        assertTrue(xml.contains("type=\"endsWith\""), "no endsWith rule");
    }

    @Test
    @DisplayName("#39 P1 a blocked dynamic array keeps its input cells after reopening")
    void blockedSpillKeepsInputs(@TempDir Path dir) throws Exception {
        Path p = dir.resolve("blocked-spill.xlsx");
        double[] written = {3, 1, 3, 2, 1};
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("S");
            for (int r = 0; r < written.length; r++) {
                s.cell(r + 1, 0).setValue(n(written[r]));
            }
            s.cell(0, 0).setFormula("SORT(UNIQUE(A2:A100))");
            wb.save(p);
        }
        try (Workbook wb = Workbook.open(p)) {
            Sheet s = wb.sheet("S");
            double[] reopened = new double[written.length];
            for (int r = 0; r < written.length; r++) {
                reopened[r] = number(s.cell(r + 1, 0).value());
            }
            assertArrayEquals(written, reopened);
        }
    }

    @Test
    @DisplayName("#40 P1 a pivot with two row fields groups each outer item once")
    void nestedPivotGroupsOuterItems() {
        try (Workbook wb = Workbook.create()) {
            Sheet d = wb.createSheet("Data");
            String[][] rows = {{"Region", "Product", "Amount"}, {"EMEA", "Widget", "10"}, {"APAC", "Widget", "20"},
                    {"EMEA", "Gadget", "30"}, {"APAC", "Gadget", "40"}, {"EMEA", "Widget", "50"}};
            for (int r = 0; r < rows.length; r++) {
                for (int c = 0; c < 3; c++) {
                    d.cell(r, c).setValue(r > 0 && c == 2 ? n(Double.parseDouble(rows[r][c])) : t(rows[r][c]));
                }
            }
            Sheet ps = wb.createSheet("Pivot");
            PivotTableView pv = ps.pivotTable(ps.addPivotTable("P", "A1", "Data!A1:C6"));
            pv.addRowField("Region");
            pv.addRowField("Product");
            pv.addDataField("Amount", AggregateFunction.SUM);
            pv.refresh();
            Range used = ps.usedRange();
            List<CellValue> labels = new ArrayList<>();
            for (int r = used.firstRow(); r <= used.lastRow(); r++) {
                labels.add(ps.cell(r, 0).value());
            }
            assertEquals(1, labels.stream().filter(t("EMEA")::equals).count(), "row labels " + labels);
            assertEquals(1, labels.stream().filter(t("APAC")::equals).count(), "row labels " + labels);
        }
    }

    @Test
    @DisplayName("#43 P1 the three-argument addTimeline is saved")
    void threeArgumentTimelineIsSaved(@TempDir Path dir) throws Exception {
        Path p = dir.resolve("timeline.xlsx");
        try (Workbook wb = Workbook.create()) {
            Sheet d = wb.createSheet("Data");
            d.cell(0, 0).setValue(t("OrderDate"));
            d.cell(0, 1).setValue(t("Amount"));
            for (int r = 1; r <= 6; r++) {
                d.cell(r, 0).setDate(LocalDate.of(2026, r, 15));
                d.cell(r, 1).setValue(n(r * 10));
            }
            Sheet ps = wb.createSheet("Pivot");
            PivotTableView pv = ps.pivotTable(ps.addPivotTable("SalesPivot", "A1", "Data!A1:B7"));
            pv.addRowField("OrderDate");
            pv.addDataField("Amount", AggregateFunction.SUM);
            pv.refresh();
            wb.addTimeline("DateTimeline", "SalesPivot", "OrderDate");
            wb.save(p);
        }
        assertTrue(zipNames(p).stream().anyMatch(name -> name.startsWith("xl/timelines/")), "no timeline part saved");
        try (Workbook wb = Workbook.open(p)) {
            assertEquals(1, wb.timelineCount());
        }
    }

    @Test
    @DisplayName("#46 P1 Theme.builder() writes no empty colour for unset slots")
    void partialThemeHasNoEmptyColours(@TempDir Path dir) throws Exception {
        Path p = dir.resolve("theme.xlsx");
        try (Workbook wb = Workbook.create()) {
            wb.createSheet("S").cell(0, 0).setValue(t("x"));
            wb.setTheme(Theme.builder().color(Theme.Slot.ACCENT1, Color.parse("#2C3E50")).majorFont("Georgia").build());
            wb.save(p);
        }
        assertFalse(zipEntry(p, "xl/theme/theme1.xml").contains("<a:srgbClr val=\"\"/>"));
    }

    @Test
    @DisplayName("#47 P1 colorScale refuses one or four colours")
    void colorScaleNeedsTwoOrThreeColours() {
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("S");
            Range range = s.range("A1:A5");
            assertThrows(IllegalArgumentException.class,
                    () -> range.addConditionalFormat(ConditionalFormat.builder().colorScale(Color.RED).build()));
            assertThrows(IllegalArgumentException.class, () -> range.addConditionalFormat(ConditionalFormat.builder()
                    .colorScale(Color.RED, Color.YELLOW, Color.GREEN, Color.BLUE).build()));
        }
    }

    @Test
    @DisplayName("#48 P1 setIterativeCalc is saved")
    void iterativeCalcIsSaved(@TempDir Path dir) throws Exception {
        Path p = dir.resolve("iterative.xlsx");
        try (Workbook wb = Workbook.create()) {
            wb.createSheet("S").cell(0, 0).setValue(n(1));
            wb.setIterativeCalc(true, 50, 0.01);
            wb.save(p);
        }
        String calcPr = findAll(zipEntry(p, "xl/workbook.xml"), "<calcPr[^>]*>").get(0);
        assertTrue(calcPr.contains("iterate=\"1\""), calcPr);
        assertTrue(calcPr.contains("iterateCount=\"50\""), calcPr);
    }

    @Test
    @DisplayName("#49 P1 createTable refuses a name that is a cell reference or already used")
    void createTableRefusesBadNames() {
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("S");
            for (int r = 0; r < 4; r++) {
                for (int c = 0; c < 5; c++) {
                    s.cell(r, c).setValue(r == 0 ? t("H" + c) : n(r));
                }
            }
            assertThrows(IllegalArgumentException.class, () -> s.range("A1:B4").createTable("T1", "H0", "H1"));
            s.range("A1:B4").createTable("Sales", "H0", "H1");
            assertThrows(IllegalArgumentException.class, () -> s.range("D1:E4").createTable("Sales", "H3", "H4"));
        }
    }

    @Test
    @DisplayName("#51 P1 a JSON round trip keeps the column order")
    void jsonKeepsColumnOrder() {
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("S");
            s.cell(0, 0).setValue(t("Region"));
            s.cell(0, 1).setValue(t("Amount"));
            s.cell(1, 0).setValue(t("APAC"));
            s.cell(1, 1).setValue(n(5));
            byte[] json = wb.saveBytes("json");
            try (Workbook back = Workbook.openBytes(json, "json")) {
                Sheet bs = back.sheet(back.sheetCount() - 1);
                assertEquals(t("Region"), bs.cell(0, 0).value(), new String(json, StandardCharsets.UTF_8));
                assertEquals(t("Amount"), bs.cell(0, 1).value());
            }
        }
    }

    @Test
    @DisplayName("#52 P1 PDF permissions without an owner password are written or refused")
    void pdfPermissionsAreNotDropped() {
        try (Workbook wb = Workbook.create()) {
            wb.createSheet("S").cell(0, 0).setValue(t("x"));
            byte[] pdf;
            try {
                pdf = Renderer.toPdf(wb, PdfOptions.builder().allowCopy(false).allowPrint(false).build());
            } catch (IllegalArgumentException | IllegalStateException refused) {
                return;
            }
            assertTrue(new String(pdf, StandardCharsets.ISO_8859_1).contains("/Encrypt"),
                    "allowCopy(false)/allowPrint(false) were dropped without an error");
        }
    }

    // ---- P2: fix when convenient ----

    @Test
    @DisplayName("#42 P2 a circular reference written before iteration is enabled converges after recalculate()")
    void iterationAfterWritingCircularFormulas() {
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("S");
            s.cell(0, 0).setFormula("B1+1");
            s.cell(0, 1).setFormula("A1/2");
            wb.setIterativeCalc(true, 100, 0.001);
            wb.recalculate();
            assertEquals(2.0, number(s.cell(0, 0).value()), 0.01);
            assertEquals(1.0, number(s.cell(0, 1).value()), 0.01);
        }
    }

    @Test
    @DisplayName("#44 P2 saving as .xlsb keeps comments and hyperlinks")
    void xlsbKeepsCommentsAndLinks() {
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("S");
            s.cell(0, 0).setValue(t("note here"));
            s.cell(0, 0).setComment("Ann", "check this");
            s.cell(1, 0).setValue(t("link"));
            s.cell(1, 0).setHyperlink("https://keikai.io");
            byte[] xlsb = wb.saveBytes("xlsb");
            try (Workbook back = Workbook.openBytes(xlsb, "xlsb")) {
                Sheet bs = back.sheet("S");
                assertNotNull(bs.cell(0, 0).comment(), "comment lost");
                assertNotNull(bs.cell(1, 0).hyperlink(), "hyperlink lost");
            }
        }
    }

    @Test
    @DisplayName("#45 P2 insertRows and insertColumns move pictures, charts and shapes")
    void insertingMovesDrawings() {
        byte[] png = Base64.getDecoder().decode(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==");
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("S");
            for (int r = 0; r < 10; r++) {
                s.cell(r, 1).setValue(n(r));
            }
            s.addPicture(png, "png", 12, 4, 14, 6);
            s.addChart(Chart.column().series("S", "S!$B$1:$B$10").anchor(12, 8, 20, 14).build());
            s.addShape("rect", "box", 12, 1, 14, 3);
            s.insertRows(0, 3);
            s.insertColumns(0, 2);
            assertArrayEquals(new int[] {15, 6, 17, 8}, s.pictureAnchor(0), "picture");
            assertArrayEquals(new int[] {15, 10, 23, 16}, s.chart(0).anchor(), "chart");
            assertArrayEquals(new int[] {15, 3, 17, 5}, s.shapeAnchor(0), "shape");
        }
    }

    @Test
    @DisplayName("#50 P2 saveEncrypted to .ods writes an ODS package or refuses")
    void encryptedOdsIsOds(@TempDir Path dir) throws Exception {
        Path p = dir.resolve("secret.ods");
        try (Workbook wb = Workbook.create()) {
            wb.createSheet("S").cell(0, 0).setValue(t("secret"));
            try {
                wb.saveEncrypted(p, "pw");
            } catch (IllegalArgumentException refused) {
                return;
            }
        }
        byte[] b = Files.readAllBytes(p);
        assertArrayEquals(new byte[] {'P', 'K', 3, 4}, Arrays.copyOf(b, 4), "not a ZIP (ODS) package");
    }

    @Test
    @DisplayName("#55 P2 evaluation mode: StreamWorkbook keeps the program's last row")
    void streamNoticeDoesNotOverwriteLastRow(@TempDir Path dir) throws Exception {
        assumeFalse(License.status().isLicensed(), "needs evaluation mode");
        Path p = dir.resolve("stream.xlsx");
        try (StreamWorkbook sw = StreamWorkbook.create(p)) {
            sw.startSheet("S");
            sw.row(0).addText("first").commit();
            sw.row(1_048_575).addText("last row").commit();
            sw.finish();
        }
        assertEquals(1, count(sheetsXml(p), "<row r=\"1048576\""), "two rows with the same number");
        try (Workbook wb = Workbook.open(p)) {
            assertEquals(t("last row"), wb.sheet("S").cell(1_048_575, 0).value());
        }
    }

    // ---- P3: minor ----

    @Test
    @DisplayName("#53 P3 barShape on a 2-D bar chart reads back as set")
    void barShapeIsSaved(@TempDir Path dir) throws Exception {
        Path p = dir.resolve("bars.xlsx");
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("S");
            for (int r = 0; r < 4; r++) {
                s.cell(r, 0).setValue(n(r + 1));
            }
            s.addChart(Chart.bar().barShape("cylinder").series("A", "S!$A$1:$A$4").anchor(0, 2, 10, 8).build());
            wb.save(p);
        }
        try (Workbook wb = Workbook.open(p)) {
            assertEquals("cylinder", wb.sheet("S").chart(0).barShape());
        }
    }

    @Test
    @DisplayName("#54 P3 conditional formats added without a priority get distinct priorities")
    void conditionalFormatPrioritiesAreDistinct(@TempDir Path dir) throws Exception {
        Path p = dir.resolve("priorities.xlsx");
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("S");
            for (int i = 0; i < 5; i++) {
                s.cell(i, 0).setValue(n(i));
            }
            s.range("A1:A5").addConditionalFormat(ConditionalFormat.builder().cellIs("greaterThan", "2").fillColor(Color.RED).build());
            s.range("A1:A5").addConditionalFormat(ConditionalFormat.builder().cellIs("lessThan", "4").fillColor(Color.GREEN).build());
            s.range("A1:A5").addConditionalFormat(ConditionalFormat.builder().dataBar(Color.BLUE).build());
            wb.save(p);
        }
        List<String> priorities = findAll(sheetsXml(p), "priority=\"\\d+\"");
        assertEquals(3, new HashSet<>(priorities).size(), "saved " + priorities);
    }

    @Test
    @DisplayName("#56 P3 consolidate(SUM) does not write -0.0 over a text label column")
    void consolidateLeavesLabelsAlone() {
        try (Workbook wb = Workbook.create()) {
            Sheet jan = wb.createSheet("Jan");
            Sheet feb = wb.createSheet("Feb");
            Sheet sum = wb.createSheet("Sum");
            for (Sheet s : new Sheet[] {jan, feb}) {
                s.cell(0, 0).setValue(t("North"));
                s.cell(0, 1).setValue(n(10));
                s.cell(1, 0).setValue(t("South"));
                s.cell(1, 1).setValue(n(20));
            }
            sum.consolidate(0, 0, ConsolidateFunction.SUM, jan.range("A1:B2"), feb.range("A1:B2"));
            CellValue label = sum.cell(0, 0).value();
            assertFalse(label instanceof CellValue.Number num && num.value() == 0
                    && 1 / num.value() < 0, "label column holds -0.0");
            assertEquals(20.0, number(sum.cell(0, 1).value()));
        }
    }

    @Test
    @DisplayName("#57 P3 addName stores the formula without a leading '='")
    void addNameStripsEquals(@TempDir Path dir) throws Exception {
        Path p = dir.resolve("names.xlsx");
        try (Workbook wb = Workbook.create()) {
            wb.createSheet("S").cell(0, 0).setFormula("100*Rate");
            wb.addName("Rate", "=0.05");
            assertEquals("0.05", wb.names().get(0).formula());
            wb.save(p);
        }
        assertTrue(zipEntry(p, "xl/workbook.xml").contains("<definedName name=\"Rate\">0.05</definedName>"));
    }

    @Test
    @DisplayName("#58 P3 evaluate rejects a malformed formula like setFormula does")
    void evaluateRejectsMalformedFormula() {
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("S");
            assertThrows(AxyraFormulaException.class, () -> s.cell(0, 0).setFormula("SUM(("));
            assertThrows(AxyraFormulaException.class, () -> wb.evaluate("SUM(("));
        }
    }

    @Test
    @DisplayName("#59 P3 a numeric user function given text returns #VALUE!, not NaN")
    void userFunctionWithTextIsValueError() {
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("S");
            wb.registerFunction("DOUBLEIT", a -> a[0] * 2);
            s.cell(0, 1).setValue(t("abc"));
            s.cell(0, 0).setFormula("DOUBLEIT(B1)");
            assertInstanceOf(CellValue.Error.class, s.cell(0, 0).value());
        }
    }

    @Test
    @DisplayName("#60 P3 CellStyle has value equality")
    void cellStyleEquals() {
        CellStyle style = CellStyle.builder().bold(true).numberFormat("#,##0.00").fill(Color.parse("#FFEB9C")).build();
        CellStyle copy = CellStyle.fromJson(style.toJson());
        assertEquals(style, copy);
        assertEquals(style.hashCode(), copy.hashCode());
    }

    @Test
    @DisplayName("#61 P3 a sheet Axyra creates has a uniqueId")
    void createdSheetHasUniqueId(@TempDir Path dir) throws Exception {
        Path p = dir.resolve("ids.xlsx");
        try (Workbook wb = Workbook.create()) {
            assertNotNull(wb.createSheet("Data").uniqueId());
            wb.save(p);
        }
        try (Workbook wb = Workbook.open(p)) {
            assertNotNull(wb.sheet("Data").uniqueId());
        }
    }
}
