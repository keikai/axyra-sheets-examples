package io.keikai.axyra.examples.regression;

import io.keikai.axyra.sheets.AggregateFunction;
import io.keikai.axyra.sheets.AxyraException;
import io.keikai.axyra.sheets.CellValue;
import io.keikai.axyra.sheets.NamedRange;
import io.keikai.axyra.sheets.PivotTableView;
import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.SortKey;
import io.keikai.axyra.sheets.Workbook;
import io.keikai.axyra.sheets.calc.CalcMode;
import io.keikai.axyra.sheets.content.Chart;
import io.keikai.axyra.sheets.content.Validation;
import io.keikai.axyra.sheets.io.CsvOptions;
import io.keikai.axyra.sheets.io.FormulaPolicy;
import io.keikai.axyra.sheets.io.SaveOptions;
import io.keikai.axyra.sheets.style.CellStyle;
import io.keikai.axyra.sheets.template.TemplateMarkers;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.Callable;
import java.util.concurrent.CountDownLatch;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.Future;

import static io.keikai.axyra.examples.regression.IssueTestSupport.findAll;
import static io.keikai.axyra.examples.regression.IssueTestSupport.n;
import static io.keikai.axyra.examples.regression.IssueTestSupport.number;
import static io.keikai.axyra.examples.regression.IssueTestSupport.sheetsXml;
import static io.keikai.axyra.examples.regression.IssueTestSupport.t;
import static io.keikai.axyra.examples.regression.IssueTestSupport.zipEntry;
import static io.keikai.axyra.examples.regression.IssueTestSupport.zipNames;
import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * Regression tests for Axyra Sheets defects found while benchmarking coding agents on the library,
 * all fixed in the build this project uses. Each test asserts the behaviour Excel has; a failure means
 * the defect is back. The numbers in the display names are the items of the benchmark issue lists
 * (0930, 1013 and 1021 builds).
 */
class FixedIssuesTest {

    @Test
    @DisplayName("#1 Range.autoFill shifts relative references")
    void autoFillShiftsReferences() {
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("S");
            s.range("A1:B2").setValues(new CellValue[][] {{n(1), n(10)}, {n(2), n(20)}});
            s.cell(2, 0).setFormula("SUM(A1:A2)");
            s.range("A3").autoFill(s.range("A3:B3"));
            assertEquals("SUM(B1:B2)", s.cell(2, 1).formula());
            assertEquals(30.0, number(s.cell(2, 1).value()));
        }
    }

    @Test
    @DisplayName("#2 insertRows adjusts other sheets' formulas and workbook names")
    void insertRowsAdjustsOtherSheetsAndNames() {
        try (Workbook wb = Workbook.create()) {
            Sheet data = wb.createSheet("Data");
            Sheet summary = wb.createSheet("Summary");
            data.range("A1:A3").setValues(new CellValue[][] {{n(1)}, {n(2)}, {n(3)}});
            summary.cell(0, 0).setFormula("SUM(Data!A1:A3)");
            wb.addName("Lines", "Data!$A$1:$A$3");
            data.insertRows(2, 1);
            data.cell(2, 0).setValue(n(100));
            assertEquals("SUM(Data!A1:A4)", summary.cell(0, 0).formula());
            assertEquals(106.0, number(summary.cell(0, 0).value()));
            String lines = wb.names().stream().filter(x -> x.name().equals("Lines"))
                    .findFirst().map(NamedRange::formula).orElseThrow();
            assertEquals("Data!$A$1:$A$4", lines);
        }
    }

    @Test
    @DisplayName("#3 SUBTOTAL(109) and AGGREGATE skip rows hidden by an AutoFilter")
    void subtotalSkipsFilteredRows() {
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("Orders");
            s.range("A1:B5").setValues(new CellValue[][] {{t("Status"), t("Amount")}, {t("Open"), n(10)},
                    {t("Shipped"), n(20)}, {t("Open"), n(30)}, {t("Shipped"), n(40)}});
            s.setAutoFilter(s.range("A1:B5"));
            s.setColumnFilter(0, "Open");
            s.applyAutoFilter();
            s.cell(6, 1).setFormula("SUBTOTAL(109,B2:B5)");
            s.cell(7, 1).setFormula("AGGREGATE(9,5,B2:B5)");
            assertEquals(40.0, number(s.cell(6, 1).value()));
            assertEquals(40.0, number(s.cell(7, 1).value()));
        }
    }

    private static Workbook region(double a, double b) {
        Workbook w = Workbook.create();
        Sheet d = w.createSheet("Data");
        d.range("A1:A2").setValues(new CellValue[][] {{n(a)}, {n(b)}});
        w.addName("Target", "Data!$A$1");
        Sheet s = w.createSheet("Summary");
        s.cell(0, 0).setFormula("SUM(Data!A1:A2)");
        s.cell(1, 0).setFormula("Target*2");
        return w;
    }

    @Test
    @DisplayName("#4 combine and copySheetFrom keep formulas and names working")
    void copiedSheetsKeepReferences() {
        try (Workbook north = region(1, 2); Workbook south = region(100, 200)) {
            north.combine(south);
            Sheet copied = north.sheet("Summary (2)");
            assertEquals("SUM('Data (2)'!A1:A2)", copied.cell(0, 0).formula());
            assertEquals(300.0, number(copied.cell(0, 0).value()));
        }
        try (Workbook pack = Workbook.create(); Workbook south = region(100, 200)) {
            pack.createSheet("Overview");
            pack.copySheetFrom(south, 0, "South Data");
            Sheet s = pack.copySheetFrom(south, 1, "South Summary");
            assertEquals(300.0, number(s.cell(0, 0).value()));
            assertEquals(200.0, number(s.cell(1, 0).value()));
        }
    }

    @Test
    @DisplayName("#5 ODS keeps formulas with FormulaPolicy.KEEP")
    void odsKeepsFormulas(@TempDir Path dir) throws Exception {
        Path p = dir.resolve("budget.ods");
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("Budget");
            s.range("A1:A2").setValues(new CellValue[][] {{n(1)}, {n(2)}});
            s.cell(2, 0).setFormula("SUM(A1:A2)");
            wb.save(p, SaveOptions.create().formulaPolicy(FormulaPolicy.KEEP));
        }
        assertTrue(zipEntry(p, "content.xml").contains("table:formula"));
        try (Workbook wb = Workbook.open(p)) {
            assertEquals("SUM(A1:A2)", wb.sheet("Budget").cell(2, 0).formula());
        }
    }

    @Test
    @DisplayName("#7 importData writes Map rows with their keys as the header")
    void importDataWritesMaps() {
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("S");
            Map<String, Object> row = new LinkedHashMap<>();
            row.put("id", "A1");
            row.put("qty", 3);
            s.importData(0, 0, List.of(row));
            assertEquals(t("id"), s.cell(0, 0).value());
            assertEquals(t("qty"), s.cell(0, 1).value());
            assertEquals(t("A1"), s.cell(1, 0).value());
            assertEquals(3.0, number(s.cell(1, 1).value()));
        }
    }

    @Test
    @DisplayName("#8 a pivot table is built when the workbook is saved, without refresh()")
    void pivotIsBuiltOnSave(@TempDir Path dir) throws Exception {
        Path p = dir.resolve("pivot.xlsx");
        try (Workbook wb = Workbook.create()) {
            Sheet d = wb.createSheet("Data");
            d.range("A1:B3").setValues(new CellValue[][] {{t("Region"), t("Amount")}, {t("East"), n(10)}, {t("West"), n(20)}});
            Sheet ps = wb.createSheet("Pivot");
            PivotTableView pv = ps.pivotTable(ps.addPivotTable("Sales", "A3", "Data!$A$1:$B$3"));
            pv.addRowField("Region");
            pv.addDataField("Amount", AggregateFunction.SUM);
            wb.save(p);
        }
        try (Workbook wb = Workbook.open(p)) {
            Sheet ps = wb.sheet("Pivot");
            List<CellValue> labels = new ArrayList<>();
            for (int r = 2; r < 8; r++) {
                labels.add(ps.cell(r, 0).value());
            }
            assertTrue(labels.contains(t("East")), "pivot labels " + labels);
            assertTrue(labels.contains(t("West")), "pivot labels " + labels);
        }
    }

    @Test
    @DisplayName("#9 setNumber keeps the cell's style")
    void setNumberKeepsStyle() {
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("S");
            s.cell(0, 0).setStyle(CellStyle.builder().bold().numberFormat("#,##0.00").build());
            s.setNumber(0, 0, 1234.5);
            assertTrue(s.cell(0, 0).style().bold());
            assertEquals("#,##0.00", s.cell(0, 0).style().numberFormat());
        }
    }

    @Test
    @DisplayName("#10 Range.setFormula on several cells shifts relative references")
    void multiCellSetFormulaShifts() {
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("S");
            s.range("A1:B3").setValues(new CellValue[][] {{n(1), n(2)}, {n(3), n(4)}, {n(5), n(6)}});
            s.range("C1:C3").setFormula("A1*B1");
            assertArrayEquals(new String[][] {{"A1*B1"}, {"A2*B2"}, {"A3*B3"}}, s.range("C1:C3").formulas());
            assertEquals(30.0, number(s.cell(2, 2).value()));
        }
    }

    @Test
    @DisplayName("#11 removeDuplicates and sort both use range-relative columns")
    void removeDuplicatesIsRangeRelative() {
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("S");
            CellValue[][] people = {{t("Name"), t("Email")}, {t("Ann"), t("ann1@x")}, {t("Ann"), t("ann2@x")},
                    {t("Bob"), t("bob@x")}, {t("Bob"), t("bob@x")}};
            s.range("C1:D5").setValues(people);
            assertEquals(1, s.range("C1:D5").removeDuplicates(new int[] {1}, true));
            assertEquals(t("ann2@x"), s.cell(2, 3).value());

            Sheet s2 = wb.createSheet("S2");
            s2.range("C1:D5").setValues(people);
            s2.range("C2:D5").sort(SortKey.desc(1));
            assertEquals(t("bob@x"), s2.cell(1, 3).value());
            assertEquals(t("ann1@x"), s2.cell(4, 3).value());
        }
    }

    @Test
    @DisplayName("#12 a Sheet handle follows its tab after another sheet is inserted")
    void sheetHandleFollowsItsTab() {
        try (Workbook wb = Workbook.create()) {
            wb.createSheet("ReadMe");
            Sheet data = wb.createSheet("Data");
            wb.insertSheet(0, "Cover");
            data.cell(0, 0).setValue(t("via handle"));
            assertEquals("Data", data.name());
            assertEquals(t("via handle"), wb.sheet("Data").cell(0, 0).value());
            assertInstanceOf(CellValue.Blank.class, wb.sheet("ReadMe").cell(0, 0).value());
        }
    }

    @Test
    @DisplayName("#13 setColumnFilter hides rows at once")
    void setColumnFilterApplies() {
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("Orders");
            s.range("A1:A3").setValues(new CellValue[][] {{t("Status")}, {t("Open")}, {t("Shipped")}});
            s.setAutoFilter(s.range("A1:A3"));
            s.setColumnFilter(0, "Open");
            assertTrue(s.isRowHidden(2));
        }
    }

    @Test
    @DisplayName("#14 Validation.list writes a quoted list and keeps the drop-down arrow")
    void validationListIsQuoted(@TempDir Path dir) throws Exception {
        Path p = dir.resolve("validation.xlsx");
        try (Workbook wb = Workbook.create()) {
            wb.createSheet("S").range("A2:A10").setValidation(Validation.list("Draft", "Approved", "Rejected"));
            wb.save(p);
        }
        String rule = findAll(sheetsXml(p), "<dataValidation .*?</dataValidation>").get(0);
        assertTrue(rule.contains("<formula1>&quot;Draft,Approved,Rejected&quot;</formula1>"), rule);
        assertFalse(rule.contains("showDropDown=\"1\""), rule);
    }

    @Test
    @DisplayName("#15 Chart.Builder.showDataLabels(true) is saved")
    void chartDataLabelsAreSaved(@TempDir Path dir) throws Exception {
        Path p = dir.resolve("chart.xlsx");
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("S");
            s.range("A1:B3").setValues(new CellValue[][] {{t("M"), t("V")}, {t("Jan"), n(1)}, {t("Feb"), n(2)}});
            s.addChart(Chart.column().title("T").series("V", "S!$B$2:$B$3", "S!$A$2:$A$3")
                    .showDataLabels(true).anchor(4, 0, 12, 8).build());
            wb.save(p);
        }
        String chart = zipEntry(p, zipNames(p).stream().filter(e -> e.startsWith("xl/charts/chart")).findFirst().orElseThrow());
        assertTrue(chart.contains("<c:showVal val=\"1\"/>"), "no value labels in the saved chart");
    }

    @Test
    @DisplayName("#16 TemplateMarkers writes a LocalDate as a formatted date")
    void templateMarkersWriteDates() {
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("T");
            s.cell(0, 0).setValue(t("${d}"));
            TemplateMarkers.process(s, Map.of("d", LocalDate.of(2026, 9, 30)));
            assertEquals(46295.0, number(s.cell(0, 0).value()));
            assertEquals("2026-09-30", s.cell(0, 0).formattedText());
        }
    }

    @Test
    @DisplayName("#17 saveCsv picks the sheet and can write raw values")
    void saveCsvWritesRawValues(@TempDir Path dir) throws Exception {
        Path p = dir.resolve("payments.csv");
        try (Workbook wb = Workbook.create()) {
            Sheet a = wb.createSheet("Invoices");
            a.cell(0, 0).setValue(n(0.125));
            a.cell(0, 0).setStyle(CellStyle.builder().numberFormat("0%").build());
            a.cell(0, 1).setValue(n(-1234.5));
            a.cell(0, 1).setStyle(CellStyle.builder().numberFormat("#,##0.00;(#,##0.00)").build());
            wb.createSheet("Payments").cell(0, 0).setValue(t("payments row"));
            wb.saveCsv(p, CsvOptions.create().sheet("Invoices").values(CsvOptions.Values.RAW));
        }
        String csv = Files.readString(p);
        assertTrue(csv.startsWith("0.125,-1234.5"), csv);
        assertFalse(csv.contains("payments row"), csv);
    }

    @Test
    @DisplayName("#19 autoFitColumn measures the formatted text")
    void autoFitUsesNumberFormat() {
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("S");
            s.cell(0, 0).setValue(n(8060));
            s.cell(0, 0).setStyle(CellStyle.builder().numberFormat("#,##0.00").build());
            s.autoFitColumn(0);
            assertTrue(s.columnWidth(0) >= 8, "width " + s.columnWidth(0) + " is too narrow for '8,060.00'");
        }
    }

    @Test
    @DisplayName("#21 CalcMode.MANUAL defers recalculation and is saved")
    void manualModeDefersAndIsSaved(@TempDir Path dir) throws Exception {
        Path p = dir.resolve("manual.xlsx");
        try (Workbook wb = Workbook.create()) {
            wb.setCalcMode(CalcMode.MANUAL);
            Sheet s = wb.createSheet("S");
            s.cell(0, 0).setValue(n(1));
            s.cell(0, 1).setFormula("A1*10");
            s.cell(0, 0).setValue(n(2));
            assertEquals(10.0, number(s.cell(0, 1).value()));
            wb.save(p);
        }
        assertTrue(zipEntry(p, "xl/workbook.xml").contains("calcMode=\"manual\""));
    }

    @Test
    @DisplayName("#21 a running balance written top-down is not quadratic")
    void formulaChainWrittenTopDownIsFast() {
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("S");
            int rows = 3000;
            for (int r = 0; r < rows; r++) {
                s.cell(r, 0).setValue(n(1));
            }
            long start = System.nanoTime();
            for (int r = 0; r < rows; r++) {
                s.cell(r, 1).setFormula(r == rows - 1 ? "A" + (r + 1) : "B" + (r + 2) + "+A" + (r + 1));
            }
            long ms = (System.nanoTime() - start) / 1_000_000;
            assertEquals(3000.0, number(s.cell(0, 1).value()));
            // 0930/1013 builds took about 12 s; fixed builds take tens of milliseconds
            assertTrue(ms < 3000, "writing the chain took " + ms + " ms");
        }
    }

    @Test
    @DisplayName("#22 two threads on one workbook get 'workbook is already in use'")
    void oneWorkbookPerThread() throws Exception {
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("S");
            ExecutorService pool = Executors.newFixedThreadPool(2);
            CountDownLatch go = new CountDownLatch(1);
            Callable<String> job = () -> {
                go.await();
                try {
                    for (int i = 0; i < 200_000; i++) {
                        s.cell(i % 100, 0).setValue(n(i));
                    }
                    return "ok";
                } catch (Exception e) {
                    return e.getMessage();
                }
            };
            Future<String> a = pool.submit(job);
            Future<String> b = pool.submit(job);
            go.countDown();
            String ra = a.get();
            String rb = b.get();
            pool.shutdown();
            String failure = ra.equals("ok") ? rb : ra;
            assertTrue(failure.contains("workbook is already in use"), failure);
        }
    }

    @Test
    @DisplayName("#26 Workbook.evaluate resolves a multi-cell defined name")
    void evaluateResolvesNames() {
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("Sheet1");
            for (int r = 1; r <= 4; r++) {
                s.cell(r, 1).setValue(n(r * 10));
            }
            wb.addName("Sales", "Sheet1!$B$2:$B$5");
            assertEquals(100.0, number(wb.evaluate("SUM(Sales)")));
            assertEquals(100.0, number(wb.evaluate("SUM(Sales)", s)));
        }
    }

    @Test
    @DisplayName("#27 manual page breaks are saved and paginate")
    void pageBreaksAreSavedAndPaginate(@TempDir Path dir) throws Exception {
        Path p = dir.resolve("breaks.xlsx");
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("P");
            for (int r = 0; r < 20; r++) {
                for (int c = 0; c < 3; c++) {
                    s.cell(r, c).setValue(n(r));
                }
            }
            assertEquals(1, wb.sheetPageCount(wb.sheet("P").index()));
            s.addRowPageBreak(10);
            s.addColumnPageBreak(1);
            assertEquals(4, wb.sheetPageCount(wb.sheet("P").index()));
            wb.save(p);
        }
        String xml = sheetsXml(p);
        assertTrue(xml.contains("<rowBreaks") && xml.contains("<colBreaks"));
        try (Workbook wb = Workbook.open(p)) {
            assertArrayEquals(new int[] {10}, wb.sheet("P").rowPageBreaks());
        }
    }

    @Test
    @DisplayName("#28 goalSeek works in CalcMode.MANUAL")
    void goalSeekInManualMode() {
        try (Workbook wb = Workbook.create()) {
            Sheet g = wb.createSheet("G");
            g.cell(0, 0).setValue(n(5));
            g.cell(1, 0).setFormula("A1*3");
            wb.setCalcMode(CalcMode.MANUAL);
            assertEquals(10.0, wb.goalSeek(g.cell(1, 0), 30, g.cell(0, 0)), 1e-9);
            assertEquals(10.0, number(g.cell(0, 0).value()), 1e-9);
        }
    }

    @Test
    @DisplayName("#29 array and dynamic-array formula text is the same after reopening")
    void arrayFormulaTextSurvivesReopen(@TempDir Path dir) throws Exception {
        Path p = dir.resolve("arrays.xlsx");
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("S");
            for (int r = 0; r < 4; r++) {
                s.cell(r, 0).setValue(n(r + 1));
            }
            s.range("B1:B4").setArrayFormula("A1:A4*2");
            s.cell(0, 2).setFormula("SORT(A1:A4,1,-1)");
            wb.save(p);
        }
        try (Workbook wb = Workbook.open(p)) {
            Sheet s = wb.sheet("S");
            assertEquals("A1:A4*2", s.cell(0, 1).formula());
            assertEquals("SORT(A1:A4,1,-1)", s.cell(0, 2).formula());
        }
    }

    @Test
    @DisplayName("Range.setArrayFormula accepts a leading '=' like setFormula")
    void setArrayFormulaStripsEquals() {
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("S");
            for (int r = 0; r < 3; r++) {
                s.cell(r, 0).setValue(n(r + 1));
            }
            s.range("B1:B3").setArrayFormula("=A1:A3*2");
            assertEquals("A1:A3*2", s.cell(0, 1).formula());
            assertEquals(6.0, number(s.cell(2, 1).value()));
        }
    }

    @Test
    @DisplayName("#30 a multi-row addSparkline makes one sparkline per row")
    void sparklinesExpandPerRow(@TempDir Path dir) throws Exception {
        Path p = dir.resolve("sparklines.xlsx");
        try (Workbook wb = Workbook.create()) {
            Sheet d = wb.createSheet("Data");
            for (int r = 1; r <= 3; r++) {
                for (int c = 1; c <= 6; c++) {
                    d.cell(r, c).setValue(n(r * c));
                }
            }
            d.addSparkline("line", "B2:G4", "H2:H4");
            wb.save(p);
        }
        List<String> sparklines = findAll(sheetsXml(p), "<x14:sparkline>.*?</x14:sparkline>");
        assertEquals(List.of(
                "<x14:sparkline><xm:f>Data!B2:G2</xm:f><xm:sqref>H2</xm:sqref></x14:sparkline>",
                "<x14:sparkline><xm:f>Data!B3:G3</xm:f><xm:sqref>H3</xm:sqref></x14:sparkline>",
                "<x14:sparkline><xm:f>Data!B4:G4</xm:f><xm:sqref>H4</xm:sqref></x14:sparkline>"), sparklines);
    }

    @Test
    @DisplayName("#31 setTimelineRange is saved")
    void timelineRangeIsSaved(@TempDir Path dir) throws Exception {
        Path p = dir.resolve("timeline.xlsx");
        try (Workbook wb = Workbook.create()) {
            Sheet data = wb.createSheet("Data");
            data.cell(0, 0).setValue(t("OrderDate"));
            data.cell(0, 1).setValue(t("Region"));
            data.cell(0, 2).setValue(t("Amount"));
            for (int i = 0; i < 12; i++) {
                data.cell(i + 1, 0).setDate(LocalDate.of(2026, 1, 5).plusDays(i * 20L));
                data.cell(i + 1, 1).setValue(t(i % 2 == 0 ? "North" : "South"));
                data.cell(i + 1, 2).setValue(n(100 + i));
            }
            Sheet ps = wb.createSheet("Pivot");
            PivotTableView pv = ps.pivotTable(ps.addPivotTable("SalesPivot", "A1", "Data!$A$1:$C$13"));
            pv.addRowField("Region");
            pv.addDataField("Amount", AggregateFunction.SUM);
            pv.refresh();
            wb.addTimeline("OrderTimeline", "SalesPivot", "OrderDate", ps.index(), 1, 5, 8, 11);
            assertTrue(wb.setTimelineRange("OrderTimeline", LocalDate.of(2026, 2, 1), LocalDate.of(2026, 6, 30)));
            wb.save(p);
        }
        try (Workbook wb = Workbook.open(p)) {
            assertArrayEquals(new LocalDate[] {LocalDate.of(2026, 2, 1), LocalDate.of(2026, 6, 30)},
                    wb.timelineRange("OrderTimeline"));
        }
    }

    @Test
    @DisplayName("#32 ChartView.showDataLabels() reads back after reopening")
    void showDataLabelsReadsBack(@TempDir Path dir) throws Exception {
        Path p = dir.resolve("labels.xlsx");
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("S");
            for (int r = 0; r < 4; r++) {
                s.cell(r, 0).setValue(t("Q" + (r + 1)));
                s.cell(r, 1).setValue(n(r + 1));
            }
            s.addChart(Chart.column().title("T").series("V", "S!$B$1:$B$4", "S!$A$1:$A$4")
                    .showDataLabels(true).anchor(0, 3, 15, 10).build());
            wb.save(p);
        }
        try (Workbook wb = Workbook.open(p)) {
            assertTrue(wb.sheet("S").chart(0).showDataLabels());
        }
    }

    @Test
    @DisplayName("#34 addTextBox writes a text box")
    void addTextBoxWritesATextBox(@TempDir Path dir) throws Exception {
        Path p = dir.resolve("textbox.xlsx");
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("S");
            int i = s.addTextBox("Notes", 1, 1, 4, 5);
            assertTrue(s.isTextBox(i));
            wb.save(p);
        }
        StringBuilder drawings = new StringBuilder();
        for (String name : zipNames(p)) {
            if (name.startsWith("xl/drawings/drawing")) {
                drawings.append(zipEntry(p, name));
            }
        }
        assertTrue(drawings.toString().contains("txBox=\"1\""));
    }

    @Test
    @DisplayName("#41 a de-DE formula reads the decimal comma")
    void localizedDecimalComma() {
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("S");
            s.cell(0, 1).setValue(n(10));
            s.cell(0, 0).setFormulaLocalized("SUMME(B1;1,5)", "de-DE");
            assertEquals("SUM(B1,1.5)", s.cell(0, 0).formula());
            assertEquals(11.5, number(s.cell(0, 0).value()));
            s.cell(0, 0).setFormula("ROUND(B1/3,2)");
            assertEquals("RUNDEN(B1/3;2)", s.cell(0, 0).formulaLocalized("de-DE"));
        }
    }

    @Test
    @DisplayName("#30 a sparkline shape that does not match the location is refused")
    void mismatchedSparklineThrows() {
        try (Workbook wb = Workbook.create()) {
            Sheet d = wb.createSheet("Data");
            for (int r = 1; r <= 3; r++) {
                for (int c = 1; c <= 6; c++) {
                    d.cell(r, c).setValue(n(r * c));
                }
            }
            AxyraException e = assertThrows(AxyraException.class,
                    () -> d.addSparkline("line", "B2:G4", "J2:J3"));
            assertTrue(e.getMessage().contains("B2:G4"), e.getMessage());
        }
    }
}
