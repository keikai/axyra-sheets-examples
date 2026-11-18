package io.keikai.axyra.examples.regression;

import io.keikai.axyra.sheets.AggregateFunction;
import io.keikai.axyra.sheets.CellValue;
import io.keikai.axyra.sheets.PivotTableView;
import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.SubtotalFunction;
import io.keikai.axyra.sheets.Workbook;
import io.keikai.axyra.sheets.calc.CalcMode;
import io.keikai.axyra.sheets.content.Validation;
import io.keikai.axyra.sheets.io.CsvOptions;
import io.keikai.axyra.sheets.io.FormulaPolicy;
import io.keikai.axyra.sheets.io.PdfOptions;
import io.keikai.axyra.sheets.io.SaveOptions;
import io.keikai.axyra.sheets.io.StreamWorkbook;
import io.keikai.axyra.sheets.style.CellStyle;
import io.keikai.axyra.sheets.style.Color;
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

import static io.keikai.axyra.examples.regression.IssueTestSupport.findAll;
import static io.keikai.axyra.examples.regression.IssueTestSupport.n;
import static io.keikai.axyra.examples.regression.IssueTestSupport.number;
import static io.keikai.axyra.examples.regression.IssueTestSupport.sheetsXml;
import static io.keikai.axyra.examples.regression.IssueTestSupport.t;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * Behaviours the Axyra documentation and the axyra-sheets skill state, checked against the build.
 * Code written from the documentation relies on each of them; a failure means the text or the
 * engine needs to change.
 */
class DocumentedBehaviourTest {

    @Test
    @DisplayName("Workbook.create() has no sheets; usedRange() of an empty sheet is null")
    void emptyWorkbookAndSheet() {
        try (Workbook wb = Workbook.create()) {
            assertEquals(0, wb.sheetCount());
            assertNull(wb.createSheet("Empty").usedRange());
        }
    }

    @Test
    @DisplayName("Validation.list refuses a value that contains a comma")
    void validationListRejectsCommas() {
        assertThrows(IllegalArgumentException.class, () -> Validation.list("a,b", "c"));
    }

    @Test
    @DisplayName("showDropdown(false) writes showDropDown=\"1\", which hides the arrow; Validation.list does not")
    void showDropdownFalseHidesTheArrow(@TempDir Path dir) throws Exception {
        Path p = dir.resolve("dropdown.xlsx");
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("S");
            s.range("A1:A3").setValidation(Validation.builder().type("list").formula1("\"x,y\"").showDropdown(false).build());
            s.range("B1:B3").setValidation(Validation.list("x", "y"));
            wb.save(p);
        }
        List<String> rules = findAll(sheetsXml(p), "<dataValidation [^>]*>");
        assertTrue(rules.stream().anyMatch(r -> r.contains("A1:A3") && r.contains("showDropDown=\"1\"")), rules.toString());
        assertTrue(rules.stream().anyMatch(r -> r.contains("B1:B3") && !r.contains("showDropDown")), rules.toString());
    }

    @Test
    @DisplayName("formula() is null for a text value that starts with '='")
    void textStartingWithEqualsStaysText() {
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("S");
            s.cell(0, 0).setValue(t("=x"));
            assertNull(s.cell(0, 0).formula());
            assertEquals(t("=x"), s.cell(0, 0).value());
        }
    }

    @Test
    @DisplayName("Sheet handles are new objects: sheets().indexOf(sheet) is -1, use sheet.index()")
    void sheetHandlesHaveNoEquality() {
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("S");
            assertEquals(-1, wb.sheets().indexOf(s));
            assertEquals(0, s.index());
        }
    }

    public static class Bean {
        public String getZeta() {
            return "z";
        }

        public String getAlpha() {
            return "a";
        }
    }

    @Test
    @DisplayName("importData: bean columns are alphabetical, Map columns in first-met order, LocalDate without a date format")
    void importDataColumnOrder() {
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("S");
            s.importData(0, 0, List.of(new Bean()));
            assertEquals(t("alpha"), s.cell(0, 0).value());
            assertEquals(t("zeta"), s.cell(0, 1).value());

            Map<String, Object> first = new LinkedHashMap<>();
            first.put("b", 1);
            first.put("a", 2);
            Map<String, Object> second = new LinkedHashMap<>();
            second.put("c", 3);
            second.put("a", 4);
            List<Map<String, Object>> rows = new ArrayList<>(List.of(first, second));
            s.importData(3, 0, rows);
            assertEquals(List.of(t("b"), t("a"), t("c")),
                    List.of(s.cell(3, 0).value(), s.cell(3, 1).value(), s.cell(3, 2).value()));

            s.importData(6, 0, List.of(Map.of("d", LocalDate.of(2024, 3, 15))));
            assertEquals(45366.0, number(s.cell(7, 0).value()));
            assertEquals("General", s.cell(7, 0).style().numberFormat());
        }
    }

    @Test
    @DisplayName("Range.style() returns the top-left cell's style")
    void rangeStyleIsTopLeft() {
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("S");
            s.cell(0, 0).setStyle(CellStyle.builder().bold().build());
            assertTrue(s.range("A1:B1").style().bold());
        }
    }

    @Test
    @DisplayName("Color.toHex() throws for a theme colour")
    void toHexNeedsAnRgbColour() {
        assertFalse(Color.theme(1).isRgb());
        assertThrows(IllegalStateException.class, () -> Color.theme(1).toHex());
    }

    @Test
    @DisplayName("Workbook.theme() is null on a new workbook")
    void newWorkbookHasNoTheme() {
        try (Workbook wb = Workbook.create()) {
            assertNull(wb.theme());
        }
    }

    @Test
    @DisplayName("an API-created pivot is blank in memory until refresh()")
    void pivotIsBlankUntilRefresh() {
        try (Workbook wb = Workbook.create()) {
            Sheet data = wb.createSheet("Data");
            data.cell(0, 0).setValue(t("R"));
            data.cell(0, 1).setValue(t("V"));
            data.cell(1, 0).setValue(t("N"));
            data.cell(1, 1).setValue(n(5));
            Sheet pivot = wb.createSheet("Pv");
            PivotTableView pv = pivot.pivotTable(pivot.addPivotTable("P", "A1", "Data!$A$1:$B$2"));
            pv.addRowField("R");
            pv.addDataField("V", AggregateFunction.SUM);
            assertNull(pivot.usedRange());
            pv.refresh();
            assertEquals(t("Row Labels"), pivot.cell(0, 0).value());
        }
    }

    @Test
    @DisplayName("subtotal returns the number of rows it inserted")
    void subtotalReturnsInsertedRows() {
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("S");
            String[][] rows = {{"Cat", "Amt"}, {"a", "1"}, {"a", "2"}, {"b", "3"}};
            for (int r = 0; r < rows.length; r++) {
                s.cell(r, 0).setValue(t(rows[r][0]));
                s.cell(r, 1).setValue(r == 0 ? t(rows[r][1]) : n(Double.parseDouble(rows[r][1])));
            }
            assertEquals(3, s.subtotal(0, 3, 0, true, SubtotalFunction.SUM, 1));
        }
    }

    @Test
    @DisplayName("renderPdf throws for an unknown sheet name")
    void unknownPdfSheetThrows(@TempDir Path dir) {
        try (Workbook wb = Workbook.create()) {
            wb.createSheet("S");
            assertThrows(RuntimeException.class,
                    () -> wb.renderPdf(dir.resolve("x.pdf"), PdfOptions.builder().sheetNames("NoSuch").build()));
        }
    }

    @Test
    @DisplayName("CalcMode.MANUAL is saved to and read from XLSX")
    void manualModeRoundTrips(@TempDir Path dir) {
        Path p = dir.resolve("manual.xlsx");
        try (Workbook wb = Workbook.create()) {
            wb.createSheet("S").cell(0, 0).setValue(n(1));
            wb.setCalcMode(CalcMode.MANUAL);
            wb.save(p);
        }
        try (Workbook wb = Workbook.open(p)) {
            assertEquals(CalcMode.MANUAL, wb.calcMode());
        }
    }

    @Test
    @DisplayName("ODS writes a formula without an OpenFormula translation as its value; KEEP makes that save throw")
    void odsFormulaPolicy(@TempDir Path dir) {
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("S");
            s.cell(0, 0).setFormula("XLOOKUP(1,{1,2},{3,4})");
            s.cell(1, 0).setFormula("SUM(1,2)");
            wb.save(dir.resolve("plain.ods"));
            IllegalArgumentException e = assertThrows(IllegalArgumentException.class, () -> wb.save(
                    dir.resolve("keep.ods"), SaveOptions.create().formulaPolicy(FormulaPolicy.KEEP)));
            assertTrue(e.getMessage().contains("S!A1"), e.getMessage());
        }
        try (Workbook wb = Workbook.open(dir.resolve("plain.ods"))) {
            Sheet s = wb.sheet("S");
            assertNull(s.cell(0, 0).formula());
            assertEquals(3.0, number(s.cell(0, 0).value()));
            assertEquals("SUM(1,2)", s.cell(1, 0).formula());
        }
    }

    @Test
    @DisplayName("CSV defaults to formatted text with CRLF line ends")
    void csvDefaults(@TempDir Path dir) throws Exception {
        Path p = dir.resolve("formatted.csv");
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("S");
            s.cell(0, 0).setValue(n(1234.5));
            s.cell(0, 0).setStyle(CellStyle.builder().numberFormat("#,##0.00").build());
            wb.saveCsv(p, CsvOptions.create().sheet("S"));
            assertTrue(new String(wb.saveBytes("csv")).contains("\r\n"));
        }
        assertTrue(Files.readString(p).startsWith("\"1,234.50\""), Files.readString(p));
    }

    @Test
    @DisplayName("StreamWorkbook refuses rows out of order and drops an uncommitted row")
    void streamRowContract(@TempDir Path dir) {
        assertThrows(RuntimeException.class, () -> {
            try (StreamWorkbook sw = StreamWorkbook.create(dir.resolve("order.xlsx"))) {
                sw.startSheet("A");
                sw.row(2).addNumber(1).commit();
                sw.row(1).addNumber(2).commit();
            }
        });
        Path p = dir.resolve("uncommitted.xlsx");
        try (StreamWorkbook sw = StreamWorkbook.create(p)) {
            sw.startSheet("A");
            sw.row(0).addNumber(1).commit();
            sw.row(1).addNumber(2);
            sw.row(2).addNumber(3).commit();
            sw.finish();
        }
        try (Workbook wb = Workbook.open(p)) {
            assertInstanceOf(CellValue.Blank.class, wb.sheet("A").cell(1, 0).value());
            assertEquals(3.0, number(wb.sheet("A").cell(2, 0).value()));
        }
    }
}
