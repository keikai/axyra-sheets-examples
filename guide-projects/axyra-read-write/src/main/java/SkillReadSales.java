import io.keikai.axyra.sheets.CellValue;
import io.keikai.axyra.sheets.Range;
import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.Workbook;
import java.nio.file.Path;
import java.util.Locale;

public final class SkillReadSales {
    public static void main(String[] args) {
        Path input = Path.of(args.length == 0 ? "input/sales.xlsx" : args[0]);
        try (Workbook workbook = Workbook.open(input)) {
            Sheet sales = workbook.sheet("Sales");
            workbook.recalculate();
            Range used = sales.usedRange();
            if (used == null) {
                System.out.println("Sales worksheet is empty");
                return;
            }
            if (used.firstRow() != 0 || used.firstColumn() != 0 || used.colCount() < 3) {
                throw new IllegalArgumentException("Expected a table starting at A1");
            }
            CellValue[][] rows = used.values();
            if (!text(rows[0][0]).equals("Order ID")
                    || !text(rows[0][1]).equals("Product")
                    || !text(rows[0][2]).equals("Amount")) {
                throw new IllegalArgumentException("Unexpected Sales headers");
            }
            for (int r = 1; r < rows.length; r++) {
                CellValue[] row = rows[r];
                if (row[0] instanceof CellValue.Blank
                        && row[1] instanceof CellValue.Blank
                        && row[2] instanceof CellValue.Blank) continue;
                String id = text(row[0]);
                double amount = number(row[2]);
                if (id.equals("Total")) {
                    System.out.printf(Locale.ROOT, "Total = %.2f%n", amount);
                    return; // The fixture's Total row marks the end of business data.
                }
                System.out.printf(Locale.ROOT, "%s | %s | %.2f%n",
                        id, text(row[1]), amount);
            }
            throw new IllegalArgumentException("Expected a Total footer");
        }
    }

    private static String text(CellValue value) {
        if (value instanceof CellValue.Text text) return text.value();
        throw new IllegalArgumentException("Expected text, found " + value);
    }

    private static double number(CellValue value) {
        if (value instanceof CellValue.Number number) return number.value();
        throw new IllegalArgumentException("Expected a number, found " + value);
    }
}
