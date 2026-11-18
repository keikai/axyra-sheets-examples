import io.keikai.axyra.sheets.CellValue;
import io.keikai.axyra.sheets.Range;
import io.keikai.axyra.sheets.Sheet;
import io.keikai.axyra.sheets.SheetType;
import io.keikai.axyra.sheets.Workbook;
import java.nio.file.Path;

public final class ReadAllWorksheets {
    public static void main(String[] args) {
        Path input = Path.of(args.length == 0 ? "input/sales.xlsx" : args[0]);
        try (Workbook wb = Workbook.open(input)) {
            for (int i = 0; i < wb.sheetCount(); i++) {
                Sheet sheet = wb.sheet(i);
                if (sheet.sheetType() != SheetType.WORKSHEET) continue;
                Range used = sheet.usedRange();
                if (used == null) {
                    System.out.println("== " + sheet.name() + " (empty)");
                    continue;
                }
                System.out.printf("== %s  rows %d-%d, cols %d-%d%n",
                        sheet.name(), used.firstRow(), used.lastRow(),
                        used.firstColumn(), used.lastColumn());
                for (CellValue[] row : used.values()) {
                    System.out.println(toTsv(row));
                }
            }
        }
    }

    public static String toTsv(CellValue[] row) {
        StringBuilder line = new StringBuilder();
        for (int c = 0; c < row.length; c++) {
            if (c > 0) line.append('\t');
            // Keep one output field per cell, including leading blanks.
            line.append(text(row[c]).replace('\t', ' ')
                    .replace('\r', ' ').replace('\n', ' '));
        }
        return line.toString();
    }

    private static String text(CellValue value) {
        if (value instanceof CellValue.Number n) return String.valueOf(n.value());
        if (value instanceof CellValue.Text t) return t.value();
        if (value instanceof CellValue.Bool b) return String.valueOf(b.value());
        if (value instanceof CellValue.Error e) return e.display();
        if (value instanceof CellValue.Array) return "[array]";
        return "";
    }
}
