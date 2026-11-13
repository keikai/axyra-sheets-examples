import io.keikai.axyra.sheets.ImportOptions;
import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.Workbook;
import io.keikai.axyra.sheets.style.CellStyle;
import java.math.BigDecimal;
import java.nio.file.Path;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;

public final class DatabaseSalesReport {
    public record Sale(String region, String product,
                       LocalDate orderDate, Double amount) {}

    public static List<Sale> readSales(ResultSet rs) throws SQLException {
        List<Sale> rows = new ArrayList<>();
        while (rs.next()) {
            BigDecimal amount = rs.getBigDecimal("amount");
            rows.add(new Sale(rs.getString("region"), rs.getString("product"),
                    rs.getObject("order_date", LocalDate.class),
                    amount == null ? null : amount.doubleValue()));
        }
        return rows;
    }

    public static void export(Connection conn, String quarter, Path output)
            throws SQLException {
        List<Sale> rows;
        try (var ps = conn.prepareStatement(
                "SELECT region, product, order_date, amount FROM sales WHERE quarter = ?")) {
            ps.setString(1, quarter);
            try (ResultSet rs = ps.executeQuery()) {
                rows = readSales(rs);
            }
        }
        writeReport(rows, output);
    }

    public static void writeReport(List<Sale> rows, Path output) {
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.createSheet("Q3 Sales");
            int written = sheet.importData(0, 0, rows, ImportOptions.builder()
                    .columns("region", "product", "orderDate", "amount")
                    .headers(true)
                    .columnLabels("Region", "Product", "Order Date", "Amount")
                    .headerStyle(CellStyle.builder().bold().build())
                    .build());
            if (!rows.isEmpty()) {
                int lastRow = rows.size(); // header is row 0
                sheet.range(1, 2, lastRow, 2).setStyle(CellStyle.builder()
                        .numberFormat("yyyy-mm-dd").build());
                sheet.range(1, 3, lastRow, 3).setStyle(CellStyle.builder()
                        .numberFormat("#,##0.00").build());
            }
            for (int c = 0; c < 4; c++) sheet.setColumnWidth(c, 18);
            // Apply all formats while the workbook is open, then save once.
            wb.save(output);
            System.out.println(written + " rows including the header");
        }
    }
}
