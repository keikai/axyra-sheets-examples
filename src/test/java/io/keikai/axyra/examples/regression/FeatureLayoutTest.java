package io.keikai.axyra.examples.regression;

import io.keikai.axyra.sheets.FilterOperator;
import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.SortKey;
import io.keikai.axyra.sheets.SubtotalFunction;
import io.keikai.axyra.sheets.TotalsRowFunction;
import io.keikai.axyra.sheets.Workbook;
import io.keikai.axyra.sheets.content.ConditionalFormat;
import io.keikai.axyra.sheets.content.TableStyleInfo;
import io.keikai.axyra.sheets.content.Validation;
import io.keikai.axyra.sheets.style.Color;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.nio.file.Path;
import java.util.List;

import static io.keikai.axyra.examples.regression.IssueTestSupport.n;
import static io.keikai.axyra.examples.regression.IssueTestSupport.number;
import static io.keikai.axyra.examples.regression.IssueTestSupport.t;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * Features F14-F23 of the axyra-sheets skill test (merges, panes, rows and columns, outlines and
 * subtotals, sorting, AutoFilter, find and replace, validation, conditional formats, tables).
 */
class FeatureLayoutTest {

    @Test
    @DisplayName("F14 merge A1:D1, list merged regions, unmerge")
    void mergeCells() {
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("S");
            s.cell(0, 0).setValue(t("Title"));
            s.range("A1:D1").merge();
            assertEquals(1, s.mergedRegions().size());
            assertEquals(3, s.mergedRegions().get(0).lastColumn());
            s.range("A1:D1").unmerge();
            assertTrue(s.mergedRegions().isEmpty());
        }
    }

    @Test
    @DisplayName("F15 frozen and split panes survive save and reopen")
    void freezeAndSplit(@TempDir Path dir) {
        Path p = dir.resolve("freeze.xlsx");
        try (Workbook wb = Workbook.create()) {
            wb.createSheet("Frozen").freezePanes(1, 1);
            wb.createSheet("Split").splitPanes(5, 3);
            wb.save(p);
        }
        try (Workbook wb = Workbook.open(p)) {
            assertEquals(1, wb.sheet("Frozen").freezeRow());
            assertEquals(1, wb.sheet("Frozen").freezeColumn());
            assertEquals(5, wb.sheet("Split").splitRow());
            assertEquals(3, wb.sheet("Split").splitColumn());
        }
    }

    @Test
    @DisplayName("F16 insert and delete rows and columns, widths, heights, hidden and auto-fit")
    void rowsAndColumns(@TempDir Path dir) {
        Path p = dir.resolve("rows.xlsx");
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("S");
            s.cell(0, 0).setValue(t("Name"));
            s.cell(1, 0).setValue(n(1));
            s.cell(2, 0).setValue(n(2));
            s.insertRows(1, 2);
            s.insertColumns(0, 1);
            s.deleteRows(1, 1);
            assertEquals(n(2), s.cell(3, 1).value());
            s.setColumnWidth(0, 18);
            s.setRowHeight(0, 30);
            s.setColumnHidden(2, true);
            s.cell(0, 3).setValue(t("a fairly long piece of text for auto-fit"));
            s.autoFitColumn(3);
            assertTrue(s.columnWidth(3) > 20);
            wb.save(p);
        }
        try (Workbook wb = Workbook.open(p)) {
            Sheet s = wb.sheet("S");
            assertEquals(18.0, s.columnWidth(0), 0.01);
            assertEquals(30.0, s.rowHeight(0), 0.01);
            assertTrue(s.isColumnHidden(2));
        }
    }

    @Test
    @DisplayName("F17 grouped and collapsed rows, grouped columns, subtotals by category")
    void groupingAndSubtotals(@TempDir Path dir) {
        Path p = dir.resolve("outline.xlsx");
        try (Workbook wb = Workbook.create()) {
            Sheet outline = wb.createSheet("Outline");
            for (int r = 0; r < 7; r++) {
                outline.cell(r, 0).setValue(n(r));
            }
            outline.groupRows(1, 4);
            outline.setRowGroupCollapsed(1, 4, true);
            outline.groupColumns(1, 2);
            Sheet sales = wb.createSheet("Sales");
            String[][] rows = {{"Category", "Item", "Amount"}, {"Dairy", "Milk", "60"}, {"Dairy", "Cheese", "150"},
                    {"Fruit", "Apple", "120"}, {"Fruit", "Banana", "80"}};
            for (int r = 0; r < rows.length; r++) {
                for (int c = 0; c < 3; c++) {
                    sales.cell(r, c).setValue(r > 0 && c == 2 ? n(Double.parseDouble(rows[r][c])) : t(rows[r][c]));
                }
            }
            assertTrue(sales.subtotal(0, 4, 0, true, SubtotalFunction.SUM, 2) > 0);
            wb.save(p);
        }
        try (Workbook wb = Workbook.open(p)) {
            Sheet outline = wb.sheet("Outline");
            assertEquals(1, outline.rowOutlineLevel(1));
            assertTrue(outline.isRowHidden(1));
            assertEquals(0, outline.rowOutlineLevel(5));
            Sheet sales = wb.sheet("Sales");
            assertEquals("SUBTOTAL(9,C2:C3)", sales.cell(3, 2).formula());
            assertEquals(210.0, number(sales.cell(3, 2).value()));
        }
    }

    @Test
    @DisplayName("F18 sort rows by two keys and sort columns left to right")
    void sorting() {
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("S");
            Object[][] data = {{"Eve", 80}, {"Bob", 95}, {"Dan", 80}, {"Alice", 95}, {"Carl", 70}};
            for (int r = 0; r < data.length; r++) {
                s.cell(r, 0).setValue(t((String) data[r][0]));
                s.cell(r, 1).setValue(n((Integer) data[r][1]));
            }
            s.range("A1:B5").sort(SortKey.desc(1), SortKey.asc(0));
            assertEquals(List.of(t("Alice"), t("Bob"), t("Dan"), t("Eve"), t("Carl")),
                    List.of(s.cell(0, 0).value(), s.cell(1, 0).value(), s.cell(2, 0).value(),
                            s.cell(3, 0).value(), s.cell(4, 0).value()));

            Sheet h = wb.createSheet("H");
            String[] names = {"Delta", "Alpha", "Charlie"};
            for (int c = 0; c < names.length; c++) {
                h.cell(0, c).setValue(t(names[c]));
                h.cell(1, c).setValue(n(c));
            }
            h.range("A1:C2").sortLeftToRight(SortKey.asc(0));
            assertEquals(t("Alpha"), h.cell(0, 0).value());
            assertEquals(t("Delta"), h.cell(0, 2).value());
        }
    }

    @Test
    @DisplayName("F19 AutoFilter by values and a greater-than condition, saved and reopened")
    void autoFilter(@TempDir Path dir) {
        Path p = dir.resolve("filter.xlsx");
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("S");
            Object[][] rows = {{"Region", "Amount"}, {"North", 120}, {"South", 80}, {"East", 200}, {"North", 90}, {"South", 140}};
            for (int r = 0; r < rows.length; r++) {
                s.cell(r, 0).setValue(t((String) rows[r][0]));
                s.cell(r, 1).setValue(r == 0 ? t((String) rows[r][1]) : n((Integer) rows[r][1]));
            }
            s.setAutoFilter(s.range("A1:B6"));
            s.setColumnFilter(0, "North", "South");
            s.setColumnCustomFilter(1, FilterOperator.GREATER_THAN, "100");
            s.applyAutoFilter();
            wb.save(p);
        }
        try (Workbook wb = Workbook.open(p)) {
            Sheet s = wb.sheet("S");
            assertFalse(s.isRowHidden(1));   // North 120
            assertTrue(s.isRowHidden(2));    // South 80
            assertTrue(s.isRowHidden(3));    // East
            assertTrue(s.isRowHidden(4));    // North 90
            assertFalse(s.isRowHidden(5));   // South 140
        }
    }

    @Test
    @DisplayName("F20 find, replace and regex replace")
    void findAndReplace() {
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("S");
            s.cell(0, 0).setValue(t("Apple"));
            s.cell(1, 0).setValue(t("Fresh apple"));
            s.cell(2, 0).setValue(t("Banana"));
            s.cell(0, 1).setValue(t("ID-1001"));
            s.cell(1, 1).setValue(t("ID-2002"));
            assertEquals(2, s.find("apple").length);
            assertEquals(2, s.replaceAll("apple", "pear"));
            assertEquals(t("Fresh pear"), s.cell(1, 0).value());
            assertEquals(2, s.replaceAllRegex("ID-\\d+", "ID-X"));
            assertEquals(t("ID-X"), s.cell(1, 1).value());
        }
    }

    @Test
    @DisplayName("F21 a drop-down list and a whole-number rule with an error message, reopened")
    void dataValidation(@TempDir Path dir) {
        Path p = dir.resolve("validation.xlsx");
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("S");
            s.range("B2:B100").setValidation(Validation.list("Open", "In Progress", "Closed"));
            s.range("C2:C100").setValidation(Validation.builder().type("whole").operator("between")
                    .formula1("1").formula2("100").errorMessage("Invalid quantity", "Enter 1 to 100.").build());
            wb.save(p);
        }
        try (Workbook wb = Workbook.open(p)) {
            List<Validation> rules = wb.sheet("S").validations();
            assertEquals(2, rules.size());
            assertTrue(rules.stream().anyMatch(v -> v.type().equals("list")));
            Validation whole = rules.stream().filter(v -> v.type().equals("whole")).findFirst().orElseThrow();
            assertEquals("100", whole.formula2());
            assertEquals("Invalid quantity", whole.errorTitle());
        }
    }

    @Test
    @DisplayName("F22 cell-value rule, three-colour scale and data bar, reopened")
    void conditionalFormatting(@TempDir Path dir) {
        Path p = dir.resolve("cf.xlsx");
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("S");
            for (int r = 0; r < 10; r++) {
                s.cell(r, 0).setValue(n(r * 25));
            }
            s.range("A1:A10").addConditionalFormat(ConditionalFormat.builder().cellIs("greaterThan", "100").fillColor(Color.RED).build());
            s.range("B1:B10").addConditionalFormat(ConditionalFormat.builder().colorScale(Color.RED, Color.YELLOW, Color.GREEN).build());
            s.range("C1:C10").addConditionalFormat(ConditionalFormat.builder().dataBar(Color.BLUE).build());
            wb.save(p);
        }
        try (Workbook wb = Workbook.open(p)) {
            List<String> types = wb.sheet("S").conditionalFormats().stream().map(ConditionalFormat::type).toList();
            assertEquals(List.of("cellIs", "colorScale", "dataBar"), types);
        }
    }

    @Test
    @DisplayName("F23 a styled table with a totals row, resized, then converted to a range")
    void tables() {
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("S");
            String[] header = {"Region", "Product", "Amount"};
            for (int c = 0; c < 3; c++) {
                s.cell(0, c).setValue(t(header[c]));
            }
            for (int r = 1; r <= 6; r++) {
                s.cell(r, 0).setValue(t("R" + r));
                s.cell(r, 1).setValue(t("P" + r));
                s.cell(r, 2).setValue(n(r * 100));
            }
            s.range("A1:C5").createTable("SalesTable", header);
            assertTrue(s.setTableStyle("SalesTable", TableStyleInfo.of("TableStyleMedium2")));
            assertTrue(s.resizeTable("SalesTable", s.range("A1:C7")));
            assertTrue(s.setTableTotalsRow("SalesTable", true));
            assertTrue(s.setTableColumnTotals("SalesTable", "Amount", TotalsRowFunction.SUM));
            assertEquals("SalesTable", s.tables().get(0).name());
            assertEquals(List.of(header), s.tables().get(0).columns());
            assertTrue(s.convertTableToRange("SalesTable"));
            assertTrue(s.tables().isEmpty());
            assertEquals(t("Region"), s.cell(0, 0).value());
        }
    }
}
