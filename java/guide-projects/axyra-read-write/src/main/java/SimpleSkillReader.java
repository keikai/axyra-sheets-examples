import io.keikai.axyra.sheets.CellValue;
import io.keikai.axyra.sheets.Workbook;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

public final class SimpleSkillReader {
    public record Order(String orderId, String product, double amount) {}
    public record Report(List<Order> orders, Double total) {}

    public static Report read(Path input) {
        try (Workbook workbook = Workbook.open(input)) {
            workbook.recalculate();
            var used = workbook.sheet("Sales").usedRange();
            if (used == null) return new Report(List.of(), null);
            CellValue[][] rows = used.values();
            if (used.colCount() < 3 || !text(rows[0][0]).equals("Order ID")
                    || !text(rows[0][1]).equals("Product")
                    || !text(rows[0][2]).equals("Amount")) {
                throw new IllegalArgumentException("Unexpected order-table headers");
            }
            List<Order> orders = new ArrayList<>();
            for (int r = 1; r < rows.length; r++) {
                CellValue[] row = rows[r];
                if (row[0] instanceof CellValue.Blank && row[1] instanceof CellValue.Blank
                        && row[2] instanceof CellValue.Blank) continue;
                String id = text(row[0]);
                if (id.equals("Total")) {
                    return new Report(List.copyOf(orders), number(row[2]));
                }
                orders.add(new Order(id, text(row[1]), number(row[2])));
            }
            throw new IllegalArgumentException("Expected a Total footer");
        }
    }

    public static void main(String[] args) {
        Report report = read(Path.of(args.length == 0 ? "input/sales.xlsx" : args[0]));
        for (Order order : report.orders()) {
            System.out.printf(Locale.ROOT, "%s | %s | %.2f%n",
                    order.orderId(), order.product(), order.amount());
        }
        if (report.total() != null) {
            System.out.printf(Locale.ROOT, "Workbook total = %.2f%n", report.total());
        }
    }

    private static String text(CellValue value) {
        if (value instanceof CellValue.Text text) return text.value();
        throw new IllegalArgumentException("Expected text, found " + value);
    }

    private static double number(CellValue value) {
        if (value instanceof CellValue.Number number) return number.value();
        throw new IllegalArgumentException("Expected number, found " + value);
    }
}
