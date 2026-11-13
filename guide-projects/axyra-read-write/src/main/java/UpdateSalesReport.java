import io.keikai.axyra.sheets.CellValue;
import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.Workbook;
import java.nio.file.Path;

public final class UpdateSalesReport {
    public static void main(String[] args) {
        try (Workbook wb = Workbook.open(Path.of("input/sales.xlsx"))) {
            Sheet sales = wb.sheet("Sales");
            sales.cell(1, 2).setValue(CellValue.number(1500));
            wb.recalculate();
            System.out.println(sales.cell(3, 2).formattedText());
            wb.save(Path.of("sales-updated.xlsx"));
        }
        try (Workbook wb = Workbook.open(Path.of("sales-updated.xlsx"))) {
            if (!wb.sheet("Sales").cell(3, 2).value().equals(CellValue.number(1580))) {
                throw new AssertionError("Saved total must be 1580");
            }
        }
    }
}
