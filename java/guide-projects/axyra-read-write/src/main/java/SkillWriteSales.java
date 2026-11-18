import io.keikai.axyra.sheets.CellValue;
import io.keikai.axyra.sheets.ImportOptions;
import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.Workbook;
import io.keikai.axyra.sheets.style.CellStyle;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

public final class SkillWriteSales {
    public record Order(String orderId, String product, double amount) {}

    public static void main(String[] args) throws Exception {
        List<Order> orders = List.of(new Order("00042", "Laptop", 1250),
                new Order("00043", "Mouse", 80));
        Path output = Path.of("output/ai-sales.xlsx");
        Files.createDirectories(output.getParent());
        write(orders, output);
        System.out.println("Saved " + output);
    }

    public static void write(List<Order> orders, Path output) {
        try (Workbook workbook = Workbook.create()) {
            Sheet sales = workbook.createSheet("Sales");
            sales.importData(0, 0, orders, ImportOptions.builder()
                    .columns("orderId", "product", "amount")
                    .columnLabels("Order ID", "Product", "Amount")
                    .headers(true)
                    .headerStyle(CellStyle.builder().bold().build())
                    .build());
            int totalRow = orders.size() + 1;
            sales.cell(totalRow, 0).setValue(CellValue.text("Total"));
            if (orders.isEmpty()) {
                sales.cell(totalRow, 2).setValue(CellValue.number(0));
            } else {
                sales.cell(totalRow, 2).setFormula("SUM(C2:C" + totalRow + ")");
            }
            sales.range(1, 2, totalRow, 2).setStyle(CellStyle.builder()
                    .numberFormat("#,##0.00").build());
            for (int c = 0; c < 3; c++) sales.setColumnWidth(c, 18);
            workbook.save(output);
        }
    }
}
