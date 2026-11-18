package io.keikai.axyra.examples.regression;

import io.keikai.axyra.sheets.Cell;
import io.keikai.axyra.sheets.CellValue;
import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.Workbook;
import io.keikai.axyra.sheets.content.Hyperlink;
import io.keikai.axyra.sheets.content.RichText;
import io.keikai.axyra.sheets.style.BorderStyle;
import io.keikai.axyra.sheets.style.CellStyle;
import io.keikai.axyra.sheets.style.Color;
import io.keikai.axyra.sheets.format.Formats;
import io.keikai.axyra.sheets.style.HAlign;
import io.keikai.axyra.sheets.style.Side;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.List;

import static io.keikai.axyra.examples.regression.IssueTestSupport.n;
import static io.keikai.axyra.examples.regression.IssueTestSupport.number;
import static io.keikai.axyra.examples.regression.IssueTestSupport.t;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * Features F01-F13 of the axyra-sheets skill test (workbooks, cell values, formulas, names, styles,
 * rich text, comments, hyperlinks). Each test does what the feature task asked and checks the result
 * the way the test judge did: in memory and, where the task saves, after reopening the file.
 */
class FeatureCoreTest {

    @Test
    @DisplayName("F01 create a workbook with two sheets and typed values, save, reopen")
    void createAndSave(@TempDir Path dir) {
        Path p = dir.resolve("data.xlsx");
        try (Workbook wb = Workbook.create()) {
            Sheet data = wb.createSheet("Data");
            wb.createSheet("Summary");
            data.range("A1:C1").setValues(new CellValue[][] {{t("Hello"), n(42.5), CellValue.bool(true)}});
            wb.save(p);
        }
        try (Workbook wb = Workbook.open(p)) {
            Sheet data = wb.sheet("Data");
            assertNotNull(wb.sheet("Summary"));
            assertEquals(t("Hello"), data.cell(0, 0).value());
            assertEquals(42.5, number(data.cell(0, 1).value()));
            assertEquals(CellValue.bool(true), data.cell(0, 2).value());
        }
    }

    @Test
    @DisplayName("F02 open a saved workbook and read every row")
    void openAndRead(@TempDir Path dir) {
        Path p = dir.resolve("input.xlsx");
        try (Workbook wb = Workbook.create()) {
            wb.createSheet("Data").range(0, 0, 3, 1).setValues(new CellValue[][] {
                    {t("name"), t("qty")}, {t("Apple"), n(10)}, {t("Banana"), n(25)}, {t("Cherry"), n(7)}});
            wb.save(p);
        }
        try (Workbook wb = Workbook.open(p)) {
            CellValue[][] rows = wb.sheet("Data").range(0, 0, 3, 1).values();
            assertEquals(t("qty"), rows[0][1]);
            assertEquals(t("Cherry"), rows[3][0]);
            assertEquals(7.0, number(rows[3][1]));
        }
    }

    @Test
    @DisplayName("F03 typed values, an error formula and formatted text")
    void typedValues() {
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("S");
            s.cell(0, 0).setValue(t("Hello"));
            s.cell(1, 0).setValue(n(1234.5678));
            s.cell(2, 0).setValue(CellValue.bool(true));
            s.cell(4, 0).setFormula("1/0");
            s.cell(5, 0).setValue(n(1234567.891));
            s.cell(5, 0).setStyle(CellStyle.builder().numberFormat("#,##0.00").build());
            assertInstanceOf(CellValue.Text.class, s.cell(0, 0).value());
            assertInstanceOf(CellValue.Number.class, s.cell(1, 0).value());
            assertInstanceOf(CellValue.Bool.class, s.cell(2, 0).value());
            assertInstanceOf(CellValue.Blank.class, s.cell(3, 0).value());
            assertInstanceOf(CellValue.Error.class, s.cell(4, 0).value());
            assertEquals("1,234,567.89", s.cell(5, 0).formattedText());
        }
    }

    @Test
    @DisplayName("F04 dates and date-times read back as LocalDate and LocalDateTime")
    void datesAndTimes() {
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("S");
            s.cell(0, 0).setDate(LocalDate.of(2024, 3, 15));
            s.cell(1, 0).setDateTime(LocalDateTime.of(2024, 3, 15, 14, 30));
            assertEquals(LocalDate.of(2024, 3, 15), s.cell(0, 0).dateValue());
            assertEquals(LocalDateTime.of(2024, 3, 15, 14, 30), s.cell(1, 0).dateTimeValue());
        }
    }

    @Test
    @DisplayName("F05 bulk write and read a 100 x 3 block")
    void bulkReadWrite() {
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("S");
            CellValue[][] data = new CellValue[100][3];
            for (int r = 0; r < 100; r++) {
                for (int c = 0; c < 3; c++) {
                    data[r][c] = n(r + c);
                }
            }
            s.range(0, 0, 99, 2).setValues(data);
            CellValue[][] read = s.range(0, 0, 99, 2).values();
            double sum = 0;
            for (CellValue[] row : read) {
                sum += number(row[1]);
            }
            assertEquals(100, read.length);
            assertEquals(5050.0, sum);
        }
    }

    @Test
    @DisplayName("F06 formulas recalculate after an input changes; evaluate an ad-hoc formula")
    void formulasAndRecalculation() {
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("S");
            for (int r = 0; r < 5; r++) {
                s.cell(r, 0).setValue(n((r + 1) * 10));
            }
            s.cell(5, 0).setFormula("SUM(A1:A5)");
            s.cell(6, 0).setFormula("AVERAGE(A1:A5)");
            wb.recalculate();
            assertEquals("SUM(A1:A5)", s.cell(5, 0).formula());
            assertEquals(150.0, number(s.cell(5, 0).value()));
            assertEquals(30.0, number(s.cell(6, 0).value()));
            s.cell(0, 0).setValue(n(100));
            wb.recalculate();
            assertEquals(240.0, number(s.cell(5, 0).value()));
            assertEquals(200.0, number(wb.evaluate("MAX(A1:A5)*2", s)));
            assertInstanceOf(CellValue.Blank.class, s.cell(7, 0).value());
        }
    }

    @Test
    @DisplayName("F07 CSE array formula and spilling dynamic arrays, also after reopening")
    void arrayFormulas(@TempDir Path dir) {
        Path p = dir.resolve("arrays.xlsx");
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("S");
            double[][] ab = {{1, 10}, {5, 3}, {3, 2.5}, {4, 1.25}, {1, 7}};
            for (int r = 0; r < ab.length; r++) {
                s.cell(r + 1, 0).setValue(n(ab[r][0]));
                s.cell(r + 1, 1).setValue(n(ab[r][1]));
            }
            s.range("C2:C6").setArrayFormula("A2:A6*B2:B6");
            s.cell(1, 4).setFormula("SEQUENCE(5,1,10,10)");
            assertEquals(10.0, number(s.cell(1, 2).value()));
            assertEquals(15.0, number(s.cell(2, 2).value()));
            assertEquals("E2:E6", s.range("E2").spillRange());
            assertEquals(50.0, number(s.cell(5, 4).value()));
            wb.save(p);
        }
        try (Workbook wb = Workbook.open(p)) {
            Sheet s = wb.sheet("S");
            assertEquals("E2:E6", s.range("E2").spillRange());
            assertEquals(7.0, number(s.cell(5, 2).value()));
        }
    }

    @Test
    @DisplayName("F08 a workbook-level name used in a formula, listed and removed")
    void namedRanges() {
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("Sheet1");
            for (int r = 1; r <= 4; r++) {
                s.cell(r, 1).setValue(n(r * 75));
            }
            wb.addName("Sales", "Sheet1!$B$2:$B$5");
            s.cell(6, 1).setFormula("SUM(Sales)");
            wb.recalculate();
            assertEquals(750.0, number(s.cell(6, 1).value()));
            assertEquals(List.of("Sales"), wb.names().stream().map(x -> x.name()).toList());
            assertTrue(wb.removeName("Sales"));
            assertNull(wb.findName("Sales"));
        }
    }

    @Test
    @DisplayName("F09 header and body cell styles survive save and reopen")
    void cellStyling(@TempDir Path dir) {
        Path p = dir.resolve("styled.xlsx");
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("S");
            s.cell(0, 0).setValue(t("Header"));
            s.range("A1:C1").setStyle(CellStyle.builder().bold().fontColor(Color.rgb(255, 255, 255))
                    .fill(Color.rgb(0x1F, 0x3A, 0x5F)).border(Side.BOTTOM, BorderStyle.THIN, Color.rgb(0, 0, 0))
                    .horizontalAlign(HAlign.CENTER).build());
            s.cell(1, 2).setValue(t("Body text that wraps"));
            s.cell(1, 2).setStyle(CellStyle.builder().italic().fontColor(Color.rgb(255, 0, 0)).wrapText(true).build());
            wb.save(p);
        }
        try (Workbook wb = Workbook.open(p)) {
            CellStyle header = wb.sheet("S").cell(0, 0).style();
            CellStyle body = wb.sheet("S").cell(1, 2).style();
            assertTrue(header.bold());
            assertEquals("FF1F3A5F", header.fillColor().toHex().replace("#", ""));
            assertEquals(BorderStyle.THIN, header.borderBottom().style());
            assertEquals(HAlign.CENTER, header.horizontalAlign());
            assertTrue(body.italic());
            assertTrue(body.wrapText());
            assertEquals("FFFF0000", body.fontColor().toHex().replace("#", ""));
        }
    }

    @Test
    @DisplayName("F10 number formats on cells and the standalone format engine")
    void numberFormats() {
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("S");
            Cell currency = s.cell(0, 0);
            currency.setValue(n(1234.5));
            currency.setStyle(CellStyle.builder().numberFormat("$#,##0.00").build());
            Cell percent = s.cell(1, 0);
            percent.setValue(n(0.1234));
            percent.setStyle(CellStyle.builder().numberFormat("0.00%").build());
            Cell thousands = s.cell(2, 0);
            thousands.setValue(n(9876543.21));
            thousands.setStyle(CellStyle.builder().numberFormat("#,##0").build());
            Cell date = s.cell(3, 0);
            date.setDate(LocalDate.of(2026, 3, 15));
            date.setStyle(date.style().toBuilder().numberFormat("yyyy-mm-dd ddd").build());
            assertEquals("$1,234.50", currency.formattedText());
            assertEquals("12.34%", percent.formattedText());
            assertEquals("9,876,543", thousands.formattedText());
            assertEquals("2026-03-15 Sun", date.formattedText());
        }
        assertEquals("$1,234.50", Formats.format(1234.5, "$#,##0.00"));
        assertEquals(0.125, Formats.parse("12.5%").number());
    }

    @Test
    @DisplayName("F11 rich text runs survive save and reopen")
    void richText(@TempDir Path dir) {
        Path p = dir.resolve("rich.xlsx");
        try (Workbook wb = Workbook.create()) {
            wb.createSheet("S").cell(0, 0).setRichText(RichText.builder()
                    .run("Important", RichText.run().bold().color(Color.rgb(255, 0, 0)))
                    .run(": please review").build());
            wb.save(p);
        }
        try (Workbook wb = Workbook.open(p)) {
            List<RichText.RunInfo> runs = wb.sheet("S").cell(0, 0).richText().runs();
            assertEquals(2, runs.size());
            assertEquals("Important", runs.get(0).text());
            assertTrue(runs.get(0).bold());
            assertFalse(runs.get(1).bold());
        }
    }

    @Test
    @DisplayName("F12 a note and a threaded comment with a reply survive save and reopen")
    void comments(@TempDir Path dir) {
        Path p = dir.resolve("comments.xlsx");
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("S");
            s.cell(0, 1).setValue(n(1));
            s.cell(0, 1).setComment("Alice", "Q3 revenue, unaudited.");
            s.cell(1, 1).setValue(n(2));
            String id = s.cell(1, 1).addThreadedComment("Bob", "Is this final?");
            s.cell(1, 1).addThreadedReply(id, "Carol", "Yes, confirmed.");
            wb.save(p);
        }
        try (Workbook wb = Workbook.open(p)) {
            Sheet s = wb.sheet("S");
            assertEquals("Alice", s.cell(0, 1).comment().author());
            assertTrue(s.cell(0, 1).comment().text().contains("Q3 revenue"));
            var thread = s.cell(1, 1).threadedComments();
            assertEquals("Is this final?", thread.get(0).text());
            assertEquals("Carol", thread.get(0).replies().get(0).author());
        }
    }

    @Test
    @DisplayName("F13 external and internal hyperlinks survive save and reopen; one is removed")
    void hyperlinks(@TempDir Path dir) {
        Path p = dir.resolve("links.xlsx");
        try (Workbook wb = Workbook.create()) {
            Sheet links = wb.createSheet("Links");
            wb.createSheet("Data").cell(1, 1).setValue(n(1));
            links.cell(0, 0).setValue(t("Visit"));
            links.cell(0, 0).setHyperlink("https://www.example.com/");
            links.cell(1, 0).setValue(t("Go to Data!B2"));
            links.cell(1, 0).setHyperlink(Hyperlink.internal("Data!B2", "Go to Data!B2"));
            wb.save(p);
        }
        try (Workbook wb = Workbook.open(p)) {
            Sheet links = wb.sheet("Links");
            assertEquals("https://www.example.com/", links.cell(0, 0).hyperlink().address());
            assertTrue(links.cell(1, 0).hyperlink().address().contains("Data!B2"));
            links.cell(0, 0).removeHyperlink();
            assertNull(links.cell(0, 0).hyperlink());
            assertNotNull(links.cell(1, 0).hyperlink());
        }
    }
}
