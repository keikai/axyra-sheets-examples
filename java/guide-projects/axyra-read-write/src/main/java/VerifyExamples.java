import io.keikai.axyra.sheets.*;
import io.keikai.axyra.sheets.io.ImageOptions;
import io.keikai.axyra.sheets.style.CellStyle;
import java.nio.file.*;
import java.time.LocalDate;
import java.util.List;
import java.util.Map;
import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;

/** Run after WriteExcelFile. Checks the exact article fixture and small recipes. */
public final class VerifyExamples {
    private static void check(boolean ok, String message) {
        if (!ok) throw new AssertionError(message);
    }
    private static double number(CellValue value) {
        return ((CellValue.Number) value).value();
    }
    public static void main(String[] args) throws Exception {
        Path output = Path.of("output/sales.xlsx");
        try (Workbook wb = Workbook.open(output)) {
            Sheet s = wb.sheet("Sales");
            check(((CellValue.Text)s.cell(1, 0).value()).value().equals("00042"), "text ID");
            check(number(s.cell(1, 2).value()) == 1250, "numeric amount");
            check("SUM(C2:C3)".equals(s.cell(3, 2).formula()), "formula");
            check(number(s.cell(3, 2).value()) == 1330, "saved formula result");
            check(s.cell(1, 2).formattedText().equals("1,250.00"), "amount format");
            check(s.cell(3, 2).formattedText().equals("1,330.00"), "total format");
            check(s.cell(0, 0).style().bold(), "bold header");
            Files.write(Path.of("output/sales-preview.png"), wb.renderRangePng(
                    s.index(), "A1:C4", ImageOptions.builder().scale(2).build()));
            s.cell(1, 2).setValue(CellValue.number(1500));
            check(number(s.cell(3, 2).value()) == 1580, "dependent recalculation");
            wb.save(Path.of("output/sales-updated.xlsx"));
        }
        try (Workbook wb = Workbook.open(Path.of("output/sales-updated.xlsx"))) {
            Sheet s = wb.sheet("Sales");
            check(number(s.cell(3, 2).value()) == 1580, "saved updated result");
            check("SUM(C2:C3)".equals(s.cell(3, 2).formula()), "saved updated formula");
            check(s.cell(1, 2).formattedText().equals("1,500.00"), "saved updated format");
        }
        // A synthetic JDBC ResultSet tests the mapping without a database server.
        try (var rs = javax.sql.rowset.RowSetProvider.newFactory().createCachedRowSet()) {
            var meta = new javax.sql.rowset.RowSetMetaDataImpl();
            meta.setColumnCount(3);
            String[] names = {"order_id", "product", "amount"};
            for (int c = 1; c <= 3; c++) {
                meta.setColumnName(c, names[c - 1]);
                meta.setColumnType(c, c == 3 ? java.sql.Types.DECIMAL : java.sql.Types.VARCHAR);
            }
            rs.setMetaData(meta);
            rs.moveToInsertRow();
            rs.updateString(1, "00042"); rs.updateString(2, "Laptop");
            rs.updateBigDecimal(3, new java.math.BigDecimal("1250.00"));
            rs.insertRow(); rs.moveToCurrentRow(); rs.beforeFirst();
            var records = SourceRecipes.readSales(rs);
            check(records.size() == 1 && records.get(0).orderId().equals("00042")
                    && records.get(0).amount() == 1250, "JDBC mapping");
            rs.beforeFirst(); rs.next(); rs.updateNull(3); rs.updateRow(); rs.beforeFirst();
            boolean rejected = false;
            try { SourceRecipes.readSales(rs); }
            catch (IllegalArgumentException expected) { rejected = true; }
            check(rejected, "JDBC null amount rejection");
        }
        // Independent bytes/stream input routes shown in the reading article.
        try (Workbook wb = Workbook.openBytes(Files.readAllBytes(output), "xlsx")) {
            check(number(wb.sheet("Sales").cell(3, 2).value()) == 1330, "bytes input");
        }
        try (var in = Files.newInputStream(output); Workbook wb = Workbook.open(in, "xlsx")) {
            check(number(wb.sheet("Sales").cell(3, 2).value()) == 1330, "stream input");
        }
        try (Workbook wb = Workbook.create()) {
            Sheet s = wb.createSheet("Recipes");
            check(s.usedRange() == null, "empty used range");
            s.cell(0, 0).setDate(LocalDate.of(2026, 10, 27));
            check(s.cell(0, 0).dateValue().equals(LocalDate.of(2026, 10, 27)), "date conversion");
            String json = "[{\"orderId\":\"00042\",\"product\":\"Laptop\",\"amount\":1250},"
                    + "{\"orderId\":\"00043\",\"product\":\"Mouse\",\"amount\":80}]";
            List<WriteExcelFile.Sale> sales = new ObjectMapper().readValue(json,
                    new TypeReference<List<WriteExcelFile.Sale>>() {});
            check(sales.size() == 2 && sales.get(0).orderId().equals("00042"), "JSON records");
            s.importData(2, 0, sales, ImportOptions.builder()
                    .columns("orderId", "product", "amount").build());
            check(number(s.cell(3, 2).value()) == 1250, "JSON import");
            s.cell(7, 0).setFormula("1/0");
            check(s.cell(7, 0).value() instanceof CellValue.Error, "error value");
            s.cell(8, 0).setValue(CellValue.text("=1+1"));
            check(s.cell(8, 0).value() instanceof CellValue.Text, "formula-like text");
            s.cell(9, 0).setFormula("SUM(C4:C5)");
            check(number(s.cell(9, 0).value()) == 1330, "dynamic total recipe");
        }
        System.out.println("PASS: text IDs, numeric amounts, bold headers, number formats,");
        System.out.println("stored SUM formula, saved total 1330, edit/reopen total 1580,");
        System.out.println("byte/stream input, dates, JSON records, empty sheets, error values,");
        System.out.println("synthetic JDBC mapping, and missing JDBC amount rejection.");
        System.out.println("Scope: this sample and listed recipes; not a benchmark or all-workbook guarantee.");
    }
}
