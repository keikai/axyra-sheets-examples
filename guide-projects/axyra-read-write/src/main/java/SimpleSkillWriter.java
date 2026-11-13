import io.keikai.axyra.sheets.CellValue;
import io.keikai.axyra.sheets.ImportOptions;
import io.keikai.axyra.sheets.Workbook;
import io.keikai.axyra.sheets.style.CellStyle;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

public final class SimpleSkillWriter {
    public record Sale(String orderId, String product, double amount) {}

    public static void write(List<Sale> sales, Path output) throws Exception {
        Path parent = output.toAbsolutePath().getParent();
        Files.createDirectories(parent);
        try (Workbook workbook = Workbook.create()) {
            var sheet = workbook.createSheet("Sales");
            sheet.importData(0, 0, sales, ImportOptions.builder()
                    .columns("orderId", "product", "amount")
                    .columnLabels("Order ID", "Product", "Amount")
                    .headers(true)
                    .headerStyle(CellStyle.builder().bold().build())
                    .build());
            int totalRow = sales.size() + 1;
            sheet.cell(totalRow, 0).setValue(CellValue.text("Total"));
            if (sales.isEmpty()) sheet.cell(totalRow, 2).setValue(CellValue.number(0));
            else sheet.cell(totalRow, 2).setFormula("SUM(C2:C" + totalRow + ")");
            sheet.range(1, 2, totalRow, 2).setStyle(CellStyle.builder()
                    .numberFormat("#,##0.00").build());
            for (int c = 0; c < 3; c++) sheet.setColumnWidth(c, 18);
            workbook.save(output);
        }
    }

    public static void main(String[] args) throws Exception {
        List<Sale> sales = List.of(new Sale("00042", "Laptop", 1250),
                new Sale("00043", "Mouse", 80));
        Path output = Path.of("output/ai-sales-simple.xlsx");
        write(sales, output);
        System.out.println("Saved " + output);
    }
}
