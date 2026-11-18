import io.keikai.axyra.sheets.CellValue;
import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.Workbook;
import java.nio.file.Path;

public final class ReadExcelFile {
    public static void main(String[] args) {
        Path input = Path.of(args.length == 0 ? "input/sales.xlsx" : args[0]);
        try (Workbook wb = Workbook.open(input)) {
            Sheet sheet = wb.sheet("Sales");
            for (CellValue[] row : sheet.range("A1:C4").values()) {
                System.out.printf("%s\t%s\t%s%n",
                        text(row[0]), text(row[1]), text(row[2]));
            }
        }
    }

    private static String text(CellValue value) {
        if (value instanceof CellValue.Text t) return t.value();
        if (value instanceof CellValue.Number n) return String.valueOf(n.value());
        if (value instanceof CellValue.Bool b) return String.valueOf(b.value());
        if (value instanceof CellValue.Error e) return e.display();
        if (value instanceof CellValue.Array) return "[array]";
        return "";
    }
}
