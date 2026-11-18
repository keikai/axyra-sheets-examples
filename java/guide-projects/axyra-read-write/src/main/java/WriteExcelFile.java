import io.keikai.axyra.sheets.CellValue;
import io.keikai.axyra.sheets.ImportOptions;
import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.Workbook;
import io.keikai.axyra.sheets.style.CellStyle;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

public final class WriteExcelFile {
    public record Sale(String orderId, String product, double amount) {}

    public static void main(String[] args) throws Exception {
        Path output = Path.of("output/sales.xlsx");
        Files.createDirectories(output.getParent());
        List<Sale> sales = List.of(
                new Sale("00042", "Laptop", 1250),
                new Sale("00043", "Mouse", 80));
        try (Workbook wb = Workbook.create()) {
            Sheet sheet = wb.createSheet("Sales");
            sheet.importData(0, 0, sales, ImportOptions.builder()
                    .columns("orderId", "product", "amount")
                    .headers(true)
                    .columnLabels("Order ID", "Product", "Amount")
                    .headerStyle(CellStyle.builder().bold().build())
                    .build());
            sheet.cell(3, 0).setValue(CellValue.text("Total"));
            sheet.cell(3, 2).setFormula("SUM(C2:C3)");
            sheet.range("C2:C4").setStyle(CellStyle.builder()
                    .numberFormat("#,##0.00").build());
            for (int c = 0; c < 3; c++) sheet.setColumnWidth(c, 18);
            wb.save(output);
        }
        System.out.println("Saved " + output);
    }
}
