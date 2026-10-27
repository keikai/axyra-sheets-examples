package io.keikai.axyra.examples.regression;

import io.keikai.axyra.sheets.AggregateFunction;
import io.keikai.axyra.sheets.CellRef;
import io.keikai.axyra.sheets.CellValue;
import io.keikai.axyra.sheets.ConsolidateFunction;
import io.keikai.axyra.sheets.ImportOptions;
import io.keikai.axyra.sheets.License;
import io.keikai.axyra.sheets.PivotTableView;
import io.keikai.axyra.sheets.SignatureInfo;
import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.Workbook;
import io.keikai.axyra.sheets.calc.CalcMode;
import io.keikai.axyra.sheets.content.FormControl;
import io.keikai.axyra.sheets.content.FormControlType;
import io.keikai.axyra.sheets.formula.Functions;
import io.keikai.axyra.sheets.io.CsvOptions;
import io.keikai.axyra.sheets.io.FormulaPolicy;
import io.keikai.axyra.sheets.io.HtmlOptions;
import io.keikai.axyra.sheets.io.PdfOptions;
import io.keikai.axyra.sheets.io.SaveOptions;
import io.keikai.axyra.sheets.io.StreamStyle;
import io.keikai.axyra.sheets.io.StreamWorkbook;
import io.keikai.axyra.sheets.style.CellStyle;
import io.keikai.axyra.sheets.style.Color;
import io.keikai.axyra.sheets.style.Theme;
import io.keikai.axyra.sheets.template.TemplateMarkers;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.security.KeyStore;
import java.security.PrivateKey;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import static io.keikai.axyra.examples.regression.IssueTestSupport.n;
import static io.keikai.axyra.examples.regression.IssueTestSupport.number;
import static io.keikai.axyra.examples.regression.IssueTestSupport.t;
import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;
import static org.junit.jupiter.api.Assumptions.assumeTrue;

/**
 * Features F36-F52 of the axyra-sheets skill test (CSV, imports, streaming, templates, functions,
 * document properties, calculation, formula notation, named styles and themes, data tools, save
 * formats, signatures, slicers, form controls, split and combine). F47 (licence API) is covered by
 * {@link LicenseFallbackTest}, which needs its own JVM.
 */
class FeatureDataTest {

    @Test
    @DisplayName("F36 import CSV text, split a semicolon column, save a sheet as raw CSV")
    void csvImportAndExport(@TempDir Path dir) throws Exception {
        byte[] csv = "Product,Qty\nApple,10\n\"Banana, ripe\",25\n".getBytes(StandardCharsets.UTF_8);
        try (Workbook wb = Workbook.openBytes(csv, "csv")) {
            Sheet s = wb.sheet(wb.sheetCount() - 1);
            assertEquals(t("Banana, ripe"), s.cell(2, 0).value());
            assertEquals(25.0, number(s.cell(2, 1).value()));
        }
        Path out = dir.resolve("semicolon.csv");
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("Semicolon");
            s.cell(0, 0).setValue(t("Name;City;Score"));
            s.cell(1, 0).setValue(t("Ann;Paris;91"));
            assertEquals(3, s.textToColumns("A1:A2", ";"));
            assertEquals(91.0, number(s.cell(1, 2).value()));
            wb.saveCsv(out, CsvOptions.create().sheet("Semicolon").values(CsvOptions.Values.RAW));
        }
        assertTrue(Files.readString(out).startsWith("Name,City,Score"), Files.readString(out));
    }

    public record Item(String sku, String name, double price) {
    }

    @Test
    @DisplayName("F37 import Java records at a cell with a header row")
    void importJavaData() {
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("S");
            List<Item> items = List.of(new Item("00042", "Widget", 2.5), new Item("00043", "Gadget", 12.75));
            int rows = s.importData(1, 1, items, ImportOptions.builder().columns("sku", "name", "price")
                    .columnLabels("SKU", "Name", "Price").build());
            assertEquals(3, rows);
            assertEquals(t("SKU"), s.cell(1, 1).value());
            assertEquals(t("00042"), s.cell(2, 1).value());
            assertEquals(12.75, number(s.cell(3, 3).value()));
        }
    }

    @Test
    @DisplayName("F38 stream 100,000 rows x 5 columns with a bold header, reopen the last row")
    void streamingWrite(@TempDir Path dir) throws Exception {
        Path p = dir.resolve("big.xlsx");
        try (StreamWorkbook sw = StreamWorkbook.create(p)) {
            int bold = sw.addStyle(StreamStyle.defaults().withBold());
            sw.startSheet("Data");
            var header = sw.row(0);
            for (String h : new String[] {"ID", "Name", "Qty", "Price", "Amount"}) {
                header.addText(h, bold);
            }
            header.commit();
            for (int r = 1; r <= 100_000; r++) {
                sw.row(r).addNumber(r).addText("Item-" + r).addNumber(1).addNumber(0.5).addNumber(0.5).commit();
            }
            sw.finish();
        }
        assertTrue(Files.size(p) > 1_000_000);
        try (Workbook wb = Workbook.open(p)) {
            Sheet s = wb.sheet("Data");
            assertEquals(100_000.0, number(s.cell(100_000, 0).value()));
            assertEquals(t("Item-100000"), s.cell(100_000, 1).value());
            assertTrue(s.cell(0, 0).style().bold());
        }
    }

    @Test
    @DisplayName("F39 template markers expand a list of maps into rows")
    void templateMarkers() {
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("T");
            s.cell(0, 0).setValue(t("${items.name}"));
            s.cell(0, 1).setValue(t("${items.price}"));
            s.cell(1, 0).setValue(t("Total"));
            s.cell(1, 1).setValue(t("${items.price:sum}"));
            TemplateMarkers.process(s, Map.of("items", List.of(
                    Map.of("name", "Apple", "price", 1.25),
                    Map.of("name", "Banana", "price", 0.5),
                    Map.of("name", "Cherry", "price", 4.0))));
            assertEquals(t("Apple"), s.cell(0, 0).value());
            assertEquals(t("Cherry"), s.cell(2, 0).value());
            assertEquals(t("Total"), s.cell(3, 0).value());
            assertEquals(5.75, number(s.cell(3, 1).value()));
        }
    }

    @Test
    @DisplayName("F40 a numeric and a typed user-defined function")
    void customFunctions() {
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("S");
            wb.registerFunction("DOUBLEIT", a -> a[0] * 2);
            wb.registerTypedFunction("GREET", argv -> CellValue.text("Hello, "
                    + (argv[0] instanceof CellValue.Text name ? name.value() : "?") + "!"));
            s.cell(0, 0).setValue(n(21));
            s.cell(0, 1).setFormula("DOUBLEIT(A1)");
            s.cell(0, 2).setFormula("GREET(\"Alice\")");
            wb.recalculate();
            assertEquals(42.0, number(s.cell(0, 1).value()));
            assertEquals(t("Hello, Alice!"), s.cell(0, 2).value());
            assertEquals(100.0, number(wb.evaluate("DOUBLEIT(50)")));
        }
    }

    @Test
    @DisplayName("F41 the function catalog")
    void functionCatalog() {
        assertTrue(Functions.count() > 400, "count " + Functions.count());
        assertTrue(Functions.contains("XLOOKUP"));
        assertFalse(Functions.contains("MADEUPFUNC"));
        assertEquals("ABS", Functions.list()[0]);
    }

    @Test
    @DisplayName("F42 document and custom properties survive save and reopen")
    void documentProperties(@TempDir Path dir) {
        Path p = dir.resolve("props.xlsx");
        try (Workbook wb = Workbook.create()) {
            wb.createSheet("S").cell(0, 0).setValue(t("x"));
            wb.setTitle("Quarterly Sales Report");
            wb.setAuthor("Jane Doe");
            wb.setSubject("Q3 figures");
            wb.setKeywords("sales, quarterly");
            wb.setCustomProperty("Department", "Finance");
            wb.setCustomProperty("Budget", 125000.5);
            wb.save(p);
        }
        try (Workbook wb = Workbook.open(p)) {
            assertEquals("Quarterly Sales Report", wb.title());
            assertEquals("Jane Doe", wb.author());
            assertEquals("Q3 figures", wb.subject());
            assertEquals("sales, quarterly", wb.keywords());
            List<String> names = wb.customProperties().stream().map(x -> x.name()).toList();
            assertTrue(names.containsAll(List.of("Department", "Budget")), names.toString());
        }
    }

    @Test
    @DisplayName("F43 manual calculation, dirty recalculation, precedents, dependents, goal seek")
    void calculationControl() {
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("Calc");
            s.cell(0, 1).setValue(n(10));
            s.cell(1, 1).setValue(n(5));
            s.cell(2, 1).setFormula("B1*B2");
            s.cell(3, 1).setFormula("B3*1.1");
            wb.setCalcMode(CalcMode.MANUAL);
            s.cell(0, 1).setValue(n(20));
            assertEquals(50.0, number(s.cell(2, 1).value()));
            wb.recalculateDirty();
            assertEquals(100.0, number(s.cell(2, 1).value()));
            List<String> precedents = wb.precedents(s, 2, 1).stream().map(CellRef::a1).toList();
            assertTrue(precedents.stream().anyMatch(a -> a.endsWith("B1"))
                    && precedents.stream().anyMatch(a -> a.endsWith("B2")), precedents.toString());
            assertTrue(wb.dependents(s, 2, 1).stream().map(CellRef::a1).anyMatch(a -> a.endsWith("B4")));
            assertEquals(40.0, wb.goalSeek(s.cell(3, 1), 220, s.cell(0, 1)), 1e-6);
        }
    }

    @Test
    @DisplayName("F44 A1 to R1C1 and back with a 1-based anchor; a German localized formula")
    void formulaNotation() {
        try (Workbook wb = Workbook.create()) {
            String r1c1 = wb.formulaToR1C1("A1+$B$2+SUM(A1:A3)", 3, 3);
            assertEquals("R[-2]C[-2]+R2C2+SUM(R[-2]C[-2]:RC[-2])", r1c1);
            assertEquals("A1+$B$2+SUM(A1:A3)", wb.formulaToA1(r1c1, 3, 3));
            Sheet s = wb.createSheet("S");
            for (int r = 0; r < 3; r++) {
                s.cell(r, 0).setValue(n((r + 1) * 10));
            }
            s.cell(4, 0).setFormulaLocalized("SUMME(A1:A3)", "de-DE");
            assertEquals("SUM(A1:A3)", s.cell(4, 0).formula());
            assertEquals(60.0, number(s.cell(4, 0).value()));
        }
    }

    @Test
    @DisplayName("F45 a named style applied and updated, default font and theme, reopened")
    void namedStylesAndThemes(@TempDir Path dir) {
        Path p = dir.resolve("styles.xlsx");
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("S");
            s.cell(0, 0).setValue(t("Header"));
            wb.registerNamedStyle("ReportHeader", CellStyle.builder().bold().fontSize(12).build());
            s.range("A1:B1").applyNamedStyle("ReportHeader");
            assertTrue(wb.updateNamedStyle("ReportHeader", CellStyle.builder().bold().italic().fontSize(14).build()));
            s.range("A1:B1").applyNamedStyle("ReportHeader");
            wb.setDefaultFont(wb.defaultStyle().toBuilder().fontName("Arial").fontSize(10).build());
            Theme.Builder theme = Theme.builder().majorFont("Georgia").minorFont("Arial");
            String[] hex = {"000000", "FFFFFF", "1F3864", "E7E6E6", "1F77B4", "FF7F0E", "2CA02C", "D62728",
                    "9467BD", "8C564B", "0563C1", "954F72"};
            Theme.Slot[] slots = Theme.Slot.values();
            for (int i = 0; i < hex.length; i++) {
                theme.color(slots[i], Color.parse("#" + hex[i]));
            }
            wb.setTheme(theme.build());
            wb.save(p);
        }
        try (Workbook wb = Workbook.open(p)) {
            assertTrue(wb.hasNamedStyle("ReportHeader"));
            CellStyle header = wb.sheet("S").cell(0, 0).style();
            assertTrue(header.italic());
            assertEquals(14.0, header.fontSize());
            assertEquals("Arial", wb.defaultStyle().fontName());
            assertEquals("Georgia", wb.theme().majorFont());
        }
    }

    @Test
    @DisplayName("F46 text to columns, remove duplicates, auto-fill a series, consolidate with SUM")
    void dataTools() {
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("S");
            s.cell(0, 0).setValue(t("a,b,c"));
            s.cell(1, 0).setValue(t("d,e,f"));
            assertEquals(3, s.textToColumns("A1:A2", ","));
            assertEquals(t("f"), s.cell(1, 2).value());

            Sheet d = wb.createSheet("D");
            String[][] rows = {{"Name", "Qty"}, {"Apple", "10"}, {"Pear", "5"}, {"Apple", "10"}};
            for (int r = 0; r < rows.length; r++) {
                d.cell(r, 0).setValue(t(rows[r][0]));
                d.cell(r, 1).setValue(r == 0 ? t(rows[r][1]) : n(Double.parseDouble(rows[r][1])));
            }
            assertEquals(1, d.range("A1:B4").removeDuplicates(new int[] {0, 1}, true));

            Sheet f = wb.createSheet("F");
            f.cell(0, 0).setValue(n(1));
            f.cell(1, 0).setValue(n(2));
            f.range("A1:A2").autoFill(f.range("A1:A10"));
            assertEquals(10.0, number(f.cell(9, 0).value()));

            Sheet jan = wb.createSheet("Jan");
            Sheet feb = wb.createSheet("Feb");
            Sheet sum = wb.createSheet("Sum");
            for (Sheet m : new Sheet[] {jan, feb}) {
                m.range("A1:B2").setValues(new CellValue[][] {{n(1), n(2)}, {n(3), n(4)}});
            }
            sum.consolidate(0, 0, ConsolidateFunction.SUM, jan.range("A1:B2"), feb.range("A1:B2"));
            assertEquals(8.0, number(sum.cell(1, 1).value()));
        }
    }

    @Test
    @DisplayName("F48 the same workbook saved as xlsx, xlsm, ods, csv, tsv, html and pdf, with formula policies")
    void saveFormats(@TempDir Path dir) throws Exception {
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("Sales");
            s.range("A1:A2").setValues(new CellValue[][] {{n(1)}, {n(2)}});
            s.cell(2, 0).setFormula("SUM(A1:A2)");
            for (String ext : new String[] {"xlsx", "xlsm", "ods", "csv", "tsv"}) {
                Path p = dir.resolve("out." + ext);
                wb.save(p);
                assertTrue(Files.size(p) > 0, ext);
            }
            wb.saveHtml(dir.resolve("out.html"), HtmlOptions.create().singleFile(true));
            assertTrue(Files.readString(dir.resolve("out.html")).contains("<table"));
            wb.renderPdf(dir.resolve("out.pdf"), PdfOptions.builder().build());
            assertTrue(Files.size(dir.resolve("out.pdf")) > 0);
            wb.save(dir.resolve("values.xlsx"), SaveOptions.create().formulaPolicy(FormulaPolicy.VALUES_ONLY));
            assertThrows(IllegalArgumentException.class,
                    () -> wb.save(dir.resolve("keep.csv"), SaveOptions.create().formulaPolicy(FormulaPolicy.KEEP)));
        }
        try (Workbook wb = Workbook.open(dir.resolve("values.xlsx"))) {
            assertNull(wb.sheet("Sales").cell(2, 0).formula());
            assertEquals(3.0, number(wb.sheet("Sales").cell(2, 0).value()));
        }
    }

    @Test
    @DisplayName("F49 sign a workbook with a key generated at run time (needs a licence)")
    void digitalSignature(@TempDir Path dir) throws Exception {
        assumeTrue(License.status().isLicensed(), "signing needs a licence; the evaluation build refuses it");
        Path keystore = dir.resolve("signer.p12");
        String keytool = Path.of(System.getProperty("java.home"), "bin",
                System.getProperty("os.name").toLowerCase().contains("win") ? "keytool.exe" : "keytool").toString();
        Process process = new ProcessBuilder(keytool, "-genkeypair", "-alias", "signer", "-keyalg", "RSA",
                "-keysize", "2048", "-sigalg", "SHA256withRSA", "-dname", "CN=Axyra Test Signer", "-validity", "30",
                "-storetype", "PKCS12", "-keystore", keystore.toString(), "-storepass", "changeit",
                "-keypass", "changeit", "-noprompt").redirectErrorStream(true).start();
        assertEquals(0, process.waitFor());
        KeyStore ks = KeyStore.getInstance("PKCS12");
        try (InputStream in = Files.newInputStream(keystore)) {
            ks.load(in, "changeit".toCharArray());
        }
        PrivateKey key = (PrivateKey) ks.getKey("signer", "changeit".toCharArray());
        byte[] certificate = ks.getCertificate("signer").getEncoded();
        Path signed = dir.resolve("signed.xlsx");
        try (Workbook wb = Workbook.create()) {
            wb.createSheet("Data").cell(0, 0).setValue(n(200));
            wb.sign(signed, key.getEncoded(), certificate);
        }
        try (Workbook wb = Workbook.open(signed)) {
            SignatureInfo info = wb.signatures().get(0);
            assertEquals("Axyra Test Signer", info.signerName());
            assertEquals(SignatureInfo.Status.VALID, info.status());
        }
    }

    @Test
    @DisplayName("F50 a pivot slicer with a selection and a timeline with a range, reopened")
    void slicersAndTimelines(@TempDir Path dir) {
        Path p = dir.resolve("slicers.xlsx");
        try (Workbook wb = Workbook.create()) {
            Sheet data = wb.createSheet("Data");
            String[] header = {"OrderDate", "Region", "Amount"};
            for (int c = 0; c < 3; c++) {
                data.cell(0, c).setValue(t(header[c]));
            }
            String[] regions = {"North", "South", "East", "West"};
            for (int i = 0; i < 12; i++) {
                data.cell(i + 1, 0).setDate(LocalDate.of(2026, 1, 5).plusDays(i * 20L));
                data.cell(i + 1, 1).setValue(t(regions[i % 4]));
                data.cell(i + 1, 2).setValue(n(100 + i));
            }
            Sheet ps = wb.createSheet("Pivot");
            PivotTableView pv = ps.pivotTable(ps.addPivotTable("SalesPivot", "A1", "Data!$A$1:$C$13"));
            pv.addRowField("Region");
            pv.addDataField("Amount", AggregateFunction.SUM);
            pv.refresh();
            wb.addSlicer("RegionSlicer", "Region", "SalesPivot", ps.index(), 1, 6, 13, 10);
            wb.addTimeline("OrderTimeline", "SalesPivot", "OrderDate", ps.index(), 15, 6, 21, 12);
            assertTrue(wb.setSlicerSelection("RegionSlicer", "North", "East"));
            assertTrue(wb.setTimelineRange("OrderTimeline", LocalDate.of(2026, 2, 1), LocalDate.of(2026, 6, 30)));
            wb.save(p);
        }
        try (Workbook wb = Workbook.open(p)) {
            assertEquals("Region", wb.slicers().get(0).sourceField());
            assertEquals(List.of("East", "North"), List.of(wb.slicerSelection("RegionSlicer")).stream().sorted().toList());
            assertArrayEquals(new LocalDate[] {LocalDate.of(2026, 2, 1), LocalDate.of(2026, 6, 30)},
                    wb.timelineRange("OrderTimeline"));
        }
    }

    @Test
    @DisplayName("F51 a linked checkbox, a combo box and a button, reopened")
    void formControls(@TempDir Path dir) {
        Path p = dir.resolve("controls.xlsx");
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("S");
            for (int r = 0; r < 3; r++) {
                s.cell(r, 4).setValue(t("Item " + r));
            }
            s.addFormControl(FormControl.checkBox("AgreeCheck").text("I agree").linkedCell("$C$1").checked(true)
                    .anchor(0, 1, 1, 2).build());
            s.addFormControl(FormControl.comboBox("FruitCombo").listFillRange("$E$1:$E$3").linkedCell("$C$4")
                    .selectedIndex(2).anchor(2, 1, 3, 3).build());
            s.addFormControl(FormControl.button("RunButton").text("Run").anchor(5, 1, 7, 3).build());
            wb.save(p);
        }
        try (Workbook wb = Workbook.open(p)) {
            List<FormControl> controls = wb.sheet("S").formControls();
            assertEquals(List.of(FormControlType.CHECK_BOX, FormControlType.COMBO_BOX, FormControlType.BUTTON),
                    controls.stream().map(FormControl::type).toList());
            assertEquals("$C$1", controls.get(0).linkedCell());
            assertTrue(controls.get(0).checked());
            assertEquals("$E$1:$E$3", controls.get(1).listFillRange());
        }
    }

    @Test
    @DisplayName("F52 combine two workbooks and split one into a file per sheet")
    void splitAndCombine(@TempDir Path dir) {
        try (Workbook a = Workbook.create(); Workbook b = Workbook.create()) {
            a.createSheet("North").cell(0, 0).setValue(n(100));
            a.createSheet("South").cell(0, 0).setValue(n(200));
            b.createSheet("East").cell(0, 0).setValue(n(300));
            a.combine(b);
            List<String> names = new ArrayList<>();
            for (Sheet s : a.sheets()) {
                names.add(s.name());
            }
            assertEquals(List.of("North", "South", "East"), names);
            List<Path> files = a.splitToFiles(dir, "part_{index}_{name}.xlsx");
            assertEquals(3, files.size());
            assertEquals("part_2_East.xlsx", files.get(2).getFileName().toString());
            try (Workbook east = Workbook.open(files.get(2))) {
                assertEquals(300.0, number(east.sheet("East").cell(0, 0).value()));
            }
        }
    }
}
